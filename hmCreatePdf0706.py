import re
import gc
import sys
import ast
import uuid
import time
import pyocr
import string
import PyPDF2
import qrcode
import hashlib
import os.path
import datetime
import importlib
import configparser
from ctypes import cdll
from pymysql import connect
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import portrait
from reportlab.pdfbase.ttfonts import TTFont
from pdfminer.pdfparser import  PDFParser
from pdfminer.pdfparser import  PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal,LAParams,LTTextBox


importlib.reload(sys)
timeStart = time.time()

proDir = os.path.split(os.path.realpath(__file__))[0]
configPath = os.path.join(proDir, "HM_config.ini")
config = configparser.ConfigParser()
config.read(configPath,encoding='UTF-8')

global upload_path
global download_path
global code_Url_Str
upload_path = config.get('File_Path', 'upload_path')
download_path = config.get('File_Path', 'download_path')
code_Url_Str = config.get('QrCode_Url', 'code_Url_Str')

global dictEye
global dictHook
dictEye = ast.literal_eval(config.get('Hm_Regular', 'dictEye'))
dictHook = ast.literal_eval(config.get('Hm_Regular', 'dictHook'))

global specialGuests
specialGuests = ast.literal_eval(config.get('Special_Guests', 'specialGuests'))

hmFont = '微软雅黑'
pdfmetrics.registerFont(TTFont(hmFont, 'msyh.ttf'))
#背钩生产总表插入
hmInsertBgProduct_Str= "INSERT INTO hm_products(`hm_pd_id`, `ld_date`, `cnhk_no`, `guest_no`, `factory_no`, `description_cn`, `product_count`, `product_unit`, `ship_date`, `cloth_isok`, `cloth_ok_date`,`created_at`) \
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s,NOW())"

class hmDataHandle(object):
    def __init__(self):
        self.conn = connect(
            host='127.0.0.1', 
            port=3306, 
            db='hmpd_erp', 
            user='root', 
            password='root', 
            charset='utf8')

    def hmInsertBgProduct_sql(self,temp,data):
        cur = self.conn.cursor()
        try:
            cur.executemany(temp,data)
            self.conn.commit()
        except:
            self.conn.rollback()
        finally:
            cur.close()
    def hmInsertBgProduct_sqlOne(self,temp):
        cur = self.conn.cursor()
        try:
            cur.execute(temp)
            self.conn.commit()
        except:
            self.conn.rollback()
        finally:
            cur.close()

class HmProduct:
    '产品基类'
    pdCount = 0
    def __init__(self):
        HmProduct.pdCount += 1
        self.hm_pd_uuid = ''
        self.is_MA = False#码装
        self.ma_D = 0   #码装分母
        self.ma_U = 0   #码装分子
        self.ma_N = 0   #码装N
        self.is_E_and_H = False#钩车眼
        self.eIsSpecial = False#眼单特殊（BGLY/LY/F1/布筒）
        self.hIsSpecial = False#钩单特殊（BGLY/LY/F1/布筒）
        self.isNeedAdd = False#是否需要加数
        self.additions = 0#加数
        self.productB = 0
        self.productP = 0
        self.clothTubeEye = 0
        self.clothTubeHook = 0
        self.productSize = 0
        self.productUnit = ''
        self.colorRamk = ''
        self.productRamk = ''
        self.productECutType = ''
        self.productHCutType = ''
        self.productHookStr = '' #钩类型
        self.productEyeStr = ''  #眼类型
        self.productHookKG = '' #钩KG
        self.productEyeKG = ''  #眼KG

        self.productIsHookPressureWord = False #钩是否压字
        self.productIsEyePressureWord = False #眼是否压字
        self.productIsEyeLooseCut = False #眼是否散口切
        

def clearNullStr(textValue):
    result = ''
    result = textValue.replace('\n', '')
    result = result.replace(' ', '')
    return result


# 获取product生产单位字段类型
# 背钩部 CHE
# textValue：生产单字符串集合
# ##销售订单
def bgGetProductDp(textValue):
    result = ''
    try:
        product_matchObj = re.search(r'生产单位:.*', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            productSN = productSN.replace('生产单位:', '')
            result = productSN
    except :
        return ''
    else:
        return result


# 获取生产单编号
# EX:CN-1905-0222
# textValue：生产单字符串集合
# ##
def bgGetProductInvoicesNum(textValue):
    result = ''
    try:
        product_matchObj = re.search( r'[^主]生产单编号:.*', textValue, re.M|re.I)
        if product_matchObj:
            productIN = product_matchObj.group()
            productIN = clearNullStr(productIN)
            productIN = productIN.replace('生产单编号:', '')
            result = productIN
    except :
        return ''
    else:
        return result


# 获取产品编号/产品颜色号
# EX:J-HE-16621-EH02160
# EX:EH02160
# textValue：生产单字符串集合
# ##
def bgGetProductNumber(textValue,hmPd):
    result = ''
    try:
        product_matchObj = re.search( r'产品编号:.*', textValue, re.M|re.I)
        if product_matchObj:
            productNum = product_matchObj.group()
            productNum = clearNullStr(productNum)
            productNum = productNum.replace('产品编号:', '')
            hmColorArr = productNum.split('-')
            if (hmColorArr):
                hmPd.productColorNum = hmColorArr[-1]
            result = productNum
    except :
        return ''
    else:
        return result


# 获取产品类型/排数/B数/尺寸
# EX：IS/HHEE-3P-2B 3/4" 51X57MM
# EX：3
# EX：2
# EX：57
# textValue：生产单字符串集合
#修复换行问题python3 -m py_compile hmCreatePdf0701.py
# ##
def bgGetProductSpecification(textValue,hmPd):
    result = ''
    try:
        product_matchObj = re.search( r'产品名称:([\s\S]*)批次:', textValue, re.M|re.I)
        if product_matchObj:
            productSF = product_matchObj.group()
            productSF = clearNullStr(productSF)
            productSF = productSF.replace('产品名称', '')
            productSF = productSF.replace('批次', '')
            productSF = productSF.replace(':', '')
            result = productSF

        matchObj = re.search( r'\d+[Bb]', result, re.M|re.I)
        if matchObj:
            bCount = matchObj.group()
            bCount = bCount.replace('B', '')
            bCount = bCount.replace('b', '')
            hmPd.productB = float(bCount)

        matchObj = re.search( r'\d+[Pp]', result, re.M|re.I)
        if matchObj:
            pCount = matchObj.group()
            pCount = pCount.replace('P', '')
            pCount = pCount.replace('p', '')
            hmPd.productP = float(pCount)

        matchObj = re.search( r'[Xx]\d+', result, re.M|re.I)
        if matchObj:
            sizeCount = matchObj.group()
            sizeCount = sizeCount.replace('X', '')
            sizeCount = sizeCount.replace('x', '')
            hmPd.productSize = float(sizeCount)

    except :
        return ''
    else:
        return result


# 获取销售单编号
# EX：SO-1905-0149
# textValue：生产单字符串集合
# ##销售订单
def bgGetProductSealNum(textValue):
    result = ''
    try:
        product_matchObj = re.search( r'来源单据:.*', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            productSN = productSN.replace('来源单据', '')
            productSN = productSN.replace(':', '')
            result = productSN
    except :
        return ''
    else:
        return result


# 获取订单数量
# EX：SO-1905-0149
# textValue：825 SET
# ##销售订单
def getHmProductCount(textValue,hmPd):
    try:
        result = 0
        matchObj = re.search( r'订单数量:.*', textValue, re.M|re.I)
        if matchObj:
            productCount = matchObj.group()
            productCount = productCount.replace('\n', '')
            productCount = productCount.replace(',', '')
            hmArr = productCount.split(' ')
            if (len(hmArr) == 3):
                hmPd.productUnit = hmArr[-1]
                productNum = hmArr[-2]
                result = float(productNum)
        else:
            return 0
    except IOError:
        return 0
    else:
        return result


# 获取客人号
# HMC1444
# textValue：生产单字符串集合
# ##
def bgGetProductGuest(textValue):
    result = ''
    try:
        product_matchObj = re.search( r'客户编号: .*', textValue, re.M|re.I)
        if product_matchObj:
            productGuest = product_matchObj.group()
            productGuest = clearNullStr(productGuest)
            productGuest = productGuest.replace('客户编号', '')
            productGuest = productGuest.replace(':', '')
            result = productGuest
    except :
        return ''
    else:
        return result


# 获取product批次字段类型
# A、B
# textValue：生产单字符串集合
# ##销售订单
def bgGetProductBatch(textValue):
    result = ''
    try:
        product_matchObj = re.search( r'批次:.*', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            productSN = productSN.replace('批次:', '')
            result = productSN
    except :
        return ''
    else:
        return result


# 获取product产品钩眼布筒字段类型,clothTubeHook=0是钩车眼
# 布筒:55MM
# 布筒:22MM
# textValue：生产单字符串集合
# productObj：对象
#  ##
def bgGetProductClothTube(textValue,productObj):
    hmClothTubeEye   =  0
    hmClothTubeHook  =  0
    try:
        product_matchObj = re.finditer( r'布筒(?::|：)\s?(\d*)MM', textValue)
        if product_matchObj:

            tubeArr = []
            for match in product_matchObj:
                item = match.group()
                item_matchObj = re.search( r'(\d\d\d|\d\d|\d)', item, re.M|re.I)

                if (item_matchObj):  
                    tubeItem = item_matchObj.group()
                    tubeArr.append(tubeItem)

            #判断钩车眼数量，如果是2的话有钩有眼；如果是1的话就有可能是钩车眼
            if(len(tubeArr) == 2 ):
                hmClothTubeEye   =  tubeArr[0]
                hmClothTubeHook  =  tubeArr[1]

            elif(len(tubeArr) == 1 ):
                hmClothTubeEye   =  tubeArr[0]

            else:
                hmClothTubeEye   =  0
                hmClothTubeHook  =  0
                #眼
                #钩
        else :
            hmClothTubeEye   =  0
            hmClothTubeHook  =  0

    except :
        hmClothTubeEye   =  0
        hmClothTubeHook  =  0
    finally:
        productObj.clothTubeEye   =  int(hmClothTubeEye)
        productObj.clothTubeHook  =  int(hmClothTubeHook)



# 获取product产品钩眼切类型字段类型,clothTubeHook=0是钩车眼
# 圆角热切
# 直角热切
# textValue：生产单字符串集合
# productObj：对象
#  ##
def bgGetProductCutType(textValue,productObj):
    productECutType = ''
    productHCutType = ''
    try:
        product_matchObj = re.finditer( r'[反对四有圆圓A直N][^,，。.:：\*\(（]{2,15}切', textValue)
        if product_matchObj:

            cutArr = []
            for match in product_matchObj:

                if (match):  
                    item = match.group()
                    item = item.replace('\n', '')
                    cutArr.append(item)

            #判断钩车眼数量，如果是2的话有钩有眼；如果是1的话就是钩车眼或单眼
            if(len(cutArr) == 2 ):
                productECutType  =  cutArr[0]
                productHCutType  =  cutArr[1]

            elif(len(cutArr) == 1 and productObj.clothTubeHook == 0):
                productECutType  =  cutArr[0]
                productHCutType  =  ''
            else:
                productECutType = ''
                productHCutType = ''
                #眼
                #钩
        else :
            productObj.productECutType = ''
            productObj.productHCutType = ''

    except :
        productObj.productECutType = ''
        productObj.productHCutType = ''
    else:
        productObj.productECutType = productECutType
        productObj.productHCutType = productHCutType


# 获取product中文详细备注
# textValue：生产单字符串集合
# ##销售订单
def bgGetProductDetilRamk(textValue,hmPd):
    result = ''
    try:
        product_matchObj = re.search( r'详细说明：([\s\S]*).。\n', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            hmArr = productSN.split('.。')
            productSN = hmArr[0]
            productSN = productSN.replace('详细说明：', '')
            productSN = productSN.replace(' ', '')

            bgGetProductClothTube(productSN,hmPd)
            bgGetProductCutType(productSN,hmPd)

            bgGetEye(productSN,hmPd)
            bgGetHook(productSN,hmPd)

            colorRamk = hmArr[1]
            colorRamk = colorRamk.replace('颜色备注:', '')
            colorRamk = clearNullStr(colorRamk)
            hmPd.colorRamk = colorRamk

            productRamk = hmArr[2]
            productRamk = productRamk.replace('产品补充说明:', '')
            productRamk = clearNullStr(productRamk)
            hmPd.productRamkadd = productRamk

            result = productSN
    except :
        return ''
    else:
        return result


# 文档md5 用于不可描述的生产单防止重复匹配
# 中文描述
# fe65976af809170084acebb8d6af1fdd
# textValue：生产单字符串集合
#'PI-1906-0148J-HE-15294-KS02066A'
# ##
def bgGetPageMd5(hmPd):
    
    textValue = str(hmPd.productCasNum)+str(hmPd.productNum)+str(hmPd.productBatch)
    pdfTxt = textValue.replace(' ', '')
    try:
        md5 = hashlib.md5()
        enc = pdfTxt.encode('utf-8','strict')
        md5.update(enc)
        result = md5.hexdigest()
    except :
        result = ''
    finally:
        return result


# 获取product中文详细备注
# textValue：生产单字符串集合
# ##销售订单
def getGuestAdditions(hmPd):
    result = 0
    try:
        product_matchObj = re.search( r'\+\d+', hmPd.productRamkadd, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = productSN.replace('+', '')
            productSN = productSN.replace(' ', '')
            result = int(productSN)
        else:
            result = 50
    except :
        result = 50
    finally:
        hmPd.additions = result
        return result

# 获取product眼类型
# textValue：生产单字符串集合
# ##销售订单
def bgGetEye(textValue,hmPd):
    result = ''
    try:
        product_matchObj = re.search( r'[比练不黑红金玫尼普铜无哑][^,，++。.:：\*\(（）))]{1,15}眼', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            result = productSN
        else:
            result = ''
    except :
        result = ''
    finally:
        hmPd.productEyeStr = result
        return result

# 获取product钩类型
# textValue：生产单字符串集合
# ##销售订单
def bgGetHook(textValue,hmPd):
    result = ''
    try:
        product_matchObj = re.search( r'[比练不黑红金玫尼普铜无哑][^,，++。.:：\*\(（））)]{1,15}[钩勾]', textValue, re.M|re.I)
        if product_matchObj:
            productSN = product_matchObj.group()
            productSN = clearNullStr(productSN)
            result = productSN
        else:
            result = ''
    except :
        result = ''
    finally:
        hmPd.productHookStr = result
        return result


# 获取product钩公斤数
# textValue：生产单字符串集合
# ##销售订单
def bgGetHookEyeKg(hookEyeType,hmPd):

    result = 0
    hookEyeKg = 0
    productHEPill = 0
    hookEyeStr = ''
    productEPillSum = 0
    try:
        if hookEyeType =='E':
            hookEyeStr = hmPd.productEyeStr
            productHEPill = hmPd.productEPill
        else :
            hookEyeStr = hmPd.productHookStr
            productHEPill = hmPd.productHPill

        #如果钩眼种类找不到，粒数为0，返回空。
        if hookEyeStr == '':
            return 0
        if productHEPill == 0:
            return 0

        if hookEyeType =='E':
            hookEyeKg = dictEye[hookEyeStr]
            productEPillSum = hmPd.productEPill *  hmPd.productP
        else :
            hookEyeKg = dictHook[hookEyeStr]
            productEPillSum = hmPd.productHPill 

        result = productEPillSum / hookEyeKg 
    except :
        return 0
    else:
        return round(result,3)



##
# 特殊情况判断
# /钩车眼/压字/印字/BGLY/LY/F1/码装/需要愈多的客人加数
# 有是否有钩眼单
# 1,2,2,2,2,2,2
#
# ##
def getHmHookEyeIsSpecial(hmPdObject):

    eIsSpecial = False
    hIsSpecial = False

    try:
        #以下客人需要多加数量付
        try:
            specialGuestsNum = specialGuests[hmPdObject.productGuest]
            hmPdObject.isNeedAdd = True
        except :
            specialGuestsNum = 0
        specialGuestsNum += 0

        if hmPdObject.isNeedAdd :
            getGuestAdditions(hmPdObject)

        bgly_matchObj = re.search( r'(码装)|(YARD)|(YDS)', hmPdObject.productSf, re.M|re.I)
        if (bgly_matchObj): 
            hmPdObject.is_MA = True
            
            bgly_matchObj = re.search( r'\d*/*\d+"', hmPdObject.productSf, re.M|re.I)
            if (bgly_matchObj):
                productSN = bgly_matchObj.group()
                productSN = productSN.replace('"', '')
                hmArr = productSN.split('/')

            if (len(hmArr) == 2 ):
                hmPdObject.ma_D = int(hmArr[1])  #码装分母
                hmPdObject.ma_U = int(hmArr[0])  #码装分子
            else:
                hmPdObject.ma_D = 1
                hmPdObject.ma_U = 1
            
            if(hmPdObject.ma_D == 2 and hmPdObject.ma_U == 1):
                hmPdObject.ma_N = 57
            elif(hmPdObject.ma_D == 4 and hmPdObject.ma_U == 3):
                hmPdObject.ma_N = 48
            elif(hmPdObject.ma_D == 16 and hmPdObject.ma_U == 11):
                hmPdObject.ma_N = 52
            elif(hmPdObject.ma_D == 1 and hmPdObject.ma_U == 1):
                hmPdObject.ma_N = 36
            elif(hmPdObject.ma_D == 16 and hmPdObject.ma_U == 15):
                hmPdObject.ma_N = 38
            elif(hmPdObject.ma_D == 32 and hmPdObject.ma_U == 19):
                hmPdObject.ma_N = 60
            elif(hmPdObject.ma_D == 38 and hmPdObject.ma_U == 15):
                hmPdObject.ma_N = 38
            elif(hmPdObject.ma_D == 8 and hmPdObject.ma_U == 5):
                hmPdObject.ma_N = 58
            elif(hmPdObject.ma_D == 8 and hmPdObject.ma_U == 7):
                hmPdObject.ma_N = 41
            elif(hmPdObject.ma_D == 16 and hmPdObject.ma_U == 9):
                hmPdObject.ma_N = 64
            elif(hmPdObject.ma_D == 16 and hmPdObject.ma_U == 13):
                hmPdObject.ma_N = 44
            elif(hmPdObject.ma_D == 64 and hmPdObject.ma_U == 25):
                hmPdObject.ma_N = 72
            elif(hmPdObject.ma_D == 32 and hmPdObject.ma_U == 21):
                hmPdObject.ma_N = 55
            else:
                hmPdObject.ma_N = 0

            hmPdObject.eIsSpecial = False
            hmPdObject.is_E_and_H = False
            return True

        bgly_matchObj = re.search( r'(钩眼车)|(钩加眼车)|(钩与眼车)', hmPdObject.productRamk, re.M|re.I)
        if (bgly_matchObj): 
            hmPdObject.eIsSpecial = True
            hmPdObject.is_E_and_H = True
            return True

        #B数大于等于5
        if hmPdObject.productB >= 5:
            eIsSpecial = True
            hIsSpecial = True
        
        #排数大于等于4
        if hmPdObject.productP >= 4:
            eIsSpecial = True
            hIsSpecial = True


        #BGLY|LY|F1 和其它特殊情况
        bgly_matchObj = re.search( r'(BGLY)|(LY)|(F1)', hmPdObject.productSf, re.M|re.I)
        if (bgly_matchObj): 
            eIsSpecial = True
            hIsSpecial = True

        #眼是否散口切
        bgly_matchObj = re.search( r'(钩位压字)|(勾位压字)|(钩压)', hmPdObject.productRamk, re.M|re.I)
        if (bgly_matchObj):
            hmPdObject.productIsHookPressureWord = True 
            hIsSpecial = True

        bgly_matchObj = re.search( r'(眼位压字)|(眼压)', hmPdObject.productRamk, re.M|re.I)
        if (bgly_matchObj): 
            hmPdObject.productIsEyePressureWord = True 
            eIsSpecial = True

        if (hmPdObject.clothTubeEye >= 66): 
            eIsSpecial = True

        if (hmPdObject.clothTubeHook >= 66): 
            hIsSpecial = True

        bgly_matchObj = re.search( r'散口', hmPdObject.productECutType, re.M|re.I)
        if (bgly_matchObj):
            hmPdObject.productIsEyeLooseCut = True  
            eIsSpecial = True

        bgly_matchObj = re.search( r'散口', hmPdObject.productHCutType, re.M|re.I)
        if (bgly_matchObj): 
            hIsSpecial = True

        bgly_matchObj = re.search( r'四角圆角', hmPdObject.productECutType, re.M|re.I)
        if (bgly_matchObj):
            eIsSpecial = True

        bgly_matchObj = re.search( r'四角圆角', hmPdObject.productHCutType, re.M|re.I)
        if (bgly_matchObj): 
            hIsSpecial = True

    except :
        hmPdObject.eIsSpecial = False
        hmPdObject.hIsSpecial = False

    else:
        hmPdObject.eIsSpecial = eIsSpecial
        hmPdObject.hIsSpecial = hIsSpecial
        return True




#创建二维码图片并保存
#codeStr        用于生成二维码的字符串
def hmCreateQRImage(codeStr):
    result = ''
    try:
        qr = qrcode.QRCode(
            version = 1,
            error_correction = qrcode.constants.ERROR_CORRECT_L,
            box_size = 2.5,
            border = 1,
        )
        hmQrCodeText = code_Url_Str + codeStr
        qr.add_data(hmQrCodeText)
        qr.make(fit = True)
        imgPath = upload_path + codeStr + '.png'
        img = qr.make_image()
        img.save(imgPath)
        img.close()
        qr.clear()
        result = imgPath
        del img
        gc.collect()
    except IOError:
        return ''
    else:
        return result


def hmCreateQRCode(hmPdObject,hookEyeType,payStr):

    codeStr = hmPdObject.hm_pd_uuid
    matchPill = 0
    matchYard = 0
    imgPath = ''

    if (hookEyeType == 'E'):
        matchPill = hmPdObject.productEPill
        matchYard = hmPdObject.productEYard
        imgPath = hmCreateQRImage(codeStr)
    else:
        matchPill = hmPdObject.productHPill
        matchYard = hmPdObject.productHYard
        imgPath = upload_path + codeStr + '.png'

    #把二维码图片、HE标题文字写入PDF
    titleX = 310
    titleY = 702
    titleXX = 360
    titleYY = 692
    titleAX = 130
    titleAY = 533
    titleFontSize = 18
    pdfPath = upload_path + str(uuid.uuid1()) + '.pdf'
    c = canvas.Canvas(pdfPath)
    c.drawImage(imgPath, 30, 720)
    c.setFillColorRGB(0,0,0)
    c.setFont(psfontname=hmFont,size=titleFontSize)
    if (hookEyeType == 'E'):
        if hmPdObject.is_E_and_H:
            c.drawString(titleX,titleY,"（E+H）")
        else:
            c.drawString(titleX,titleY,"（E）")
    else:
        c.drawString(titleX,titleY,"（H）")

    c.setFont(psfontname=hmFont,size=50)
    if (hookEyeType == 'E'):
        if hmPdObject.is_E_and_H:
            c.drawString(titleX+80,titleYY,"眼钩")
        else:
            c.drawString(titleXX,titleYY,"眼")
        
        #眼压字
        if (hmPdObject.productIsEyePressureWord):
            c.drawString(titleXX,100,"压字")
        #眼散口切
        if (hmPdObject.productIsEyeLooseCut):
            c.drawString(titleXX,100,"散口")
    else:
        c.drawString(titleXX,titleYY,"钩")
        #钩压字
        if (hmPdObject.productIsHookPressureWord):
            c.drawString(titleXX,100,"压字")

    #粒数码数
    c.setStrokeColorRGB(0, 1, 0)
    c.setFillAlpha(0.4)
    c.rect(25,705,90,13,0,1)
    c.rect(25,688,90,13,0,1)
    c.setFillAlpha(1)
    
    
    c.setFont(psfontname=hmFont,size=18)
    if(hmPdObject.is_MA == False):
        c.drawString(30,705,str(matchYard) + " Y")
    else:
        c.drawString(30,705,str(int(hmPdObject.productCount)) + "+"+str(matchYard) + " Y")
    c.drawString(30,688,str(matchPill) + " PCS")
    

    #以下客人需要多加数量付
    if hmPdObject.isNeedAdd:
        csum = hmPdObject.additions + hmPdObject.productCount
        cText = "+"+str(hmPdObject.additions)+"="+str(int(csum))
        c.setFont(psfontname=hmFont,size=20)
        c.drawString(titleAX,titleAY,cText)

    c.setFillColorRGB(0.5,0.5,0.5)
    c.setFont(psfontname=hmFont,size=10)
    c.drawString(30,30,hmPdObject.productEyeStr+": "+str(hmPdObject.productEyeKG) + " kg")
    c.drawString(30,43,hmPdObject.productHookStr+": "+str(hmPdObject.productHookKG) + " kg")
    
    #工资
    c.setFont(psfontname=hmFont,size=10)
    c.drawString(25,17,payStr )
    c.save()

    del c
    gc.collect()
    
    #创建成功PDF成功移除图片
    if (hookEyeType == 'H'):
        os.remove(imgPath)
    if (hmPdObject.clothTubeHook == 0):
        os.remove(imgPath)
    return pdfPath


##
# 获取兴文计算工资粒数
# EH_car\H_car\E_car
# EH_cat\H_cat\E_cat
# ##
def getPayrollHmPillX(payrollType,hmPd):
    productSum = hmPd.productCount
    productB = hmPd.productB
    productP = hmPd.productP
    unitPrice = 0

    #以下客人需要多加50付
    if hmPd.isNeedAdd:
        productSum += hmPd.additions
    # 订单粒数【payrollP】 = （订单数【S】 * B数）
    payrollP =  productSum * productB

    #判断【衣车钩眼】订单总价
    if payrollType == 'EH_car':
        if productP <= 3:
            unitPrice = (22.1 if (payrollP <= 5000) else 17.7)
        else:
            unitPrice = (24.65 if (payrollP <= 5000) else 20.5)
    
    #判断【衣车钩】订单单价
    elif payrollType == 'H_car':
        unitPrice = (7.2 if (payrollP <= 5000) else 5.8)

    #判断【衣车眼】订单单价
    elif payrollType == 'E_car':
        if productP <= 3:
            unitPrice = (14.9 if (payrollP <= 5000) else 11.9)
        else:
            unitPrice = (17.45 if (payrollP <= 5000) else 14.7)

#-------------------------------------------------------------

    #判断【切机钩眼】订单总价
    elif payrollType == 'EH_cat':
        if productP <= 3:
            unitPrice = (12.2 if (payrollP <= 5000) else 11.4)
        else:
            unitPrice = (12.9 if (payrollP <= 5000) else 12)
    
    #判断【切机钩】订单单价
    elif payrollType == 'H_cat':
        unitPrice = (4.27 if (payrollP <= 5000) else 4.27)

    #判断【切机眼】订单单价
    elif payrollType == 'E_cat':
        if productP <= 3:
            unitPrice = (7.93 if (payrollP <= 5000) else 7.13)
        else:
            unitPrice = (8.63 if (payrollP <= 5000) else 7.73)


    if (payrollType == 'EH_car') or(payrollType == 'H_car')  or (payrollType == 'E_car'):
        qt = productB / 2
        unitPrice *= qt

    if payrollType == 'EH_cat' or payrollType == 'H_cat'  or payrollType == 'E_cat':
       if productB >= 5:
        qt = productB / 2
        unitPrice *= qt


    result =  unitPrice * (productSum/1200)

    return float('%.2f' % result)




##
# 获取兴文计算工资粒数
# EH_car\H_car\E_car
# EH_cat\H_cat\E_cat
# ##
def getPayrollHmPill(payrollType,productSum,productB,productP):
    unitPrice = 0
    # 订单粒数【payrollP】 = （订单数【S】 * B数）
    payrollP =  productSum * productB

    #判断【衣车钩眼】订单总价
    if payrollType == 'EH_car':
        if productP <= 3:
            unitPrice = (22.1 if (payrollP <= 5000) else 17.7)
        else:
            unitPrice = (24.65 if (payrollP <= 5000) else 20.5)
    
    #判断【衣车钩】订单单价
    elif payrollType == 'H_car':
        unitPrice = (7.2 if (payrollP <= 5000) else 5.8)

    #判断【衣车眼】订单单价
    elif payrollType == 'E_car':
        if productP <= 3:
            unitPrice = (14.9 if (payrollP <= 5000) else 11.9)
        else:
            unitPrice = (17.45 if (payrollP <= 5000) else 14.7)

#-------------------------------------------------------------

    #判断【切机钩眼】订单总价
    elif payrollType == 'EH_cat':
        if productP <= 3:
            unitPrice = (12.2 if (payrollP <= 5000) else 11.4)
        else:
            unitPrice = (12.9 if (payrollP <= 5000) else 12)
    
    #判断【切机钩】订单单价
    elif payrollType == 'H_cat':
        unitPrice = (4.27 if (payrollP <= 5000) else 4.27)

    #判断【切机眼】订单单价
    elif payrollType == 'E_cat':
        if productP <= 3:
            unitPrice = (7.93 if (payrollP <= 5000) else 7.13)
        else:
            unitPrice = (8.63 if (payrollP <= 5000) else 7.73)


    if (payrollType == 'EH_car') or(payrollType == 'H_car')  or (payrollType == 'E_car'):
        qt = productB / 2
        unitPrice *= qt

    if payrollType == 'EH_cat' or payrollType == 'H_cat'  or payrollType == 'E_cat':
       if productB >= 5:
        qt = productB / 2
        unitPrice *= qt


    result =  unitPrice * (productSum/1200)

    return float('%.2f' % result)
    #*数量





# 计算钩眼粒数
# （粒数）【T】= 订单数量 * 预多百分数 * B数
# 
# btType       钩眼类型  H/E
# productSum   订单总数
# productB    B数
# productP    P数
# ##
def countHmPillx(hookEyeType,hmPdObject):
    productSum = hmPdObject.productCount
    productB = hmPdObject.productB
    isSpecial = (hmPdObject.eIsSpecial if (hookEyeType == 'E') else hmPdObject.hIsSpecial) 

    if(hmPdObject.is_MA):
        productSum *= hmPdObject.ma_N

    #以下客人需要多加50付
    if hmPdObject.isNeedAdd:
        productSum += hmPdObject.additions

    try:
        result = 0
        #预多增量
        beforeCount = 1
        #初始订单数
        beginSum = productSum

        if isSpecial:
            beforeCount += 0.06
                          
        if productSum <= 500:
            beforeCount += (0.12 if (hookEyeType == 'E') else 0.06)
            beginSum += (50 if (hmPdObject.is_MA==False) else 0)

        elif 501 <= productSum and productSum <= 2000:
            beforeCount += (0.1 if (hookEyeType == 'E') else 0.06)
            beginSum += (50 if (hmPdObject.is_MA==False) else 0)
        
        elif 2001 <= productSum and productSum <= 5000 :
            beforeCount += (0.08 if (hookEyeType == 'E') else 0.06)
            beginSum += (50 if (hmPdObject.is_MA==False) else 0)
        
        elif 5001 <= productSum and productSum <= 10000 :
            beforeCount += (0.06 if (hookEyeType == 'E') else 0.05)
        
        else:
            beforeCount += (0.05 if (hookEyeType == 'E') else 0.04)

        #
        if(hmPdObject.is_MA):
            result = beginSum * beforeCount 
        else:
            result = beginSum * beforeCount * productB
            
    except :
        return 0
    else:
        return round(result)


##
# 计算布料码数
# 【M】= T粒数 / （914 / 尺寸 * B数）
# hookEyePillCount   粒数
# productSize        尺寸
# productB           B数
# 
# ##
def countHmYard(hookEyePillCount,hmPd):


    productSize = hmPd.productSize
    productB = hmPd.productB

    try:

        result = 0

        if(hmPd.is_MA == False):
            intValue = ( 914 / productSize * productB )
            intValueM = int(intValue)
            yardCount = hookEyePillCount / intValueM
        
        else:
            if(hmPd.ma_N != 0):
                yardCount = hookEyePillCount / hmPd.ma_N - hmPd.productCount
            else:
                yardCount = 0
        
        result = yardCount

    except :
        return 0
    else:
        return round(result)



def parse(fileName,productDataItemArr,isSaveInData):

    text_path = upload_path + fileName + ".pdf"
    #hmPdfSavePath = download_path + fileName + "_NEW.pdf"
    hmPdfSaveName = ""
    fileOpen = open(text_path,'rb')
    doc = PDFDocument()

    #用文件对象创建一个PDF文档分析器
    parser = PDFParser(fileOpen)
    parser.set_document(doc)
    doc.set_parser(parser)
    #提供初始化密码，如果没有密码，就创建一个空的字符串
    doc.initialize()

    #原文件
    hmPdfReaderEye = PyPDF2.PdfFileReader(fileOpen)
    hmPdfReaderHook = PyPDF2.PdfFileReader(fileOpen)
    #待写入数据文件
    hmPdfWriter = PyPDF2.PdfFileWriter()

 
    #检测文档是否提供txt转换，不提供就忽略
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        #创建PDF，资源管理器，来共享资源
        rsrcmgr = PDFResourceManager()
        #创建一个PDF设备对象
        device = PDFPageAggregator(rsrcmgr,laparams=LAParams())
        #创建一个PDF解释其对象
        interpreter = PDFPageInterpreter(rsrcmgr,device)

        openFileArr = []

        allPages = doc.get_pages()
        for page in allPages:

            interpreter.process_page(page)
            layout = device.get_result()

            textValueArr = []
            for x in layout:
                if(isinstance(x,LTTextBoxHorizontal)):
                    textValueArr.append(x.get_text())
            pdfTxt = ''.join(textValueArr)
            textValueArr.clear()
            
            dpText = bgGetProductDp(pdfTxt)
            if dpText != 'CHE':
                #print('不是背钩部生产单')
                continue

            #init
            hmPd = HmProduct()
            #生产车间
            hmPd.productDp = dpText
            #生产单编号
            hmPd.productCasNum = bgGetProductInvoicesNum(pdfTxt)
            #产品编号
            hmPd.productNum = bgGetProductNumber(pdfTxt,hmPd)
            #钩眼规格/排数/B数/尺寸
            hmPd.productSf = bgGetProductSpecification(pdfTxt,hmPd)
            #销售单号
            hmPd.productSealNum = bgGetProductSealNum(pdfTxt)
            #订单数量/单位
            hmPd.productCount = getHmProductCount(pdfTxt,hmPd)
            #客人号
            hmPd.productGuest  = bgGetProductGuest(pdfTxt)
            #产品批次
            hmPd.productBatch = bgGetProductBatch(pdfTxt)
            #产品批次
            hmPd.productBatch = bgGetProductBatch(pdfTxt)
            #产品中文描述/颜色备注/产品补充说明/钩眼布筒/钩眼切法
            hmPd.productRamk = bgGetProductDetilRamk(pdfTxt,hmPd)
            #生成生产单uuid
            hmPd.hm_pd_uuid = bgGetPageMd5(hmPd)
            #处理特殊订单
            getHmHookEyeIsSpecial(hmPd)


            #产品排数粒数
            hmPd.productEPill = countHmPillx('E',hmPd)
            hmPd.productHPill = countHmPillx('H',hmPd)
            hmPd.productEYard = countHmYard(hmPd.productEPill,hmPd)
            hmPd.productHYard = countHmYard(hmPd.productHPill,hmPd)

            #产品钩眼公斤数
            hmPd.productEyeKG = bgGetHookEyeKg('E',hmPd)
            hmPd.productHookKG = bgGetHookEyeKg('H',hmPd)
            
            hmPd.productEH_car = getPayrollHmPillX('EH_car',hmPd)
            hmPd.productH_car = getPayrollHmPillX('H_car',hmPd)
            hmPd.productE_car = getPayrollHmPillX('E_car',hmPd)

            hmPd.productEH_cat = getPayrollHmPillX('EH_cat',hmPd)
            hmPd.productH_cat = getPayrollHmPillX('H_cat',hmPd)
            hmPd.productE_cat = getPayrollHmPillX('E_cat',hmPd)


            #----------------1、生成文件_STAR----------------#
            if isSaveInData != 1:
                #工资字符串
                payStr = ''.join([
                    "（车|钩眼:",
                    str(hmPd.productEH_car),
                    "、钩:",
                    str(hmPd.productH_car),
                    "、眼:",
                    str(hmPd.productE_car),
                    "）    （切|钩眼:",
                    str(hmPd.productEH_cat),
                    "、钩:",
                    str(hmPd.productH_cat),
                    "、眼:",
                    str(hmPd.productE_cat),
                    ")"
                ])

                layoutPageId = layout.pageid - 1 



                #生成【眼】单
                eyeNewPage = hmPdfReaderEye.getPage(layoutPageId) 
                ePatch = hmCreateQRCode(hmPd,'E',payStr)
                eMarkFile = open(ePatch,'rb')
                pdfECodePage = PyPDF2.PdfFileReader(eMarkFile)
                eyeNewPage.mergePage(pdfECodePage.getPage(0))
                hmPdfWriter.addPage(eyeNewPage)
                openFileArr.append(eMarkFile)
               # os.remove(ePatch)
                del eyeNewPage
                del pdfECodePage
                gc.collect()
                #生成【钩】单
                if hmPd.clothTubeHook > 0:
                    hookNewPage = hmPdfReaderHook.getPage(layoutPageId) 
                    hPatch = hmCreateQRCode(hmPd,'H',payStr)
                    hMarkFile = open(hPatch,'rb')
                    pdfHCodePage = PyPDF2.PdfFileReader(hMarkFile)
                    hookNewPage.mergePage(pdfHCodePage.getPage(0))
                    hmPdfWriter.addPage(hookNewPage)
                    openFileArr.append(hMarkFile)
                   # os.remove(hPatch)
                    del hookNewPage
                    del pdfHCodePage
                    gc.collect()
                #用销售单号做文件名
                if hmPdfSaveName == "":
                    hmPdfSaveName = hmPd.productSealNum
            #----------------1、生成文件_END----------------#


            #----------------2、写入字典用来插入数据库_STAR----------------#
            elif isSaveInData == 1:
                itemStep = (
                    hmPd.hm_pd_uuid,
                    "2019-06-11",
                    hmPd.productSealNum,
                    hmPd.productGuest,
                    hmPd.productSf,
                    hmPd.productRamk,
                    hmPd.productCount,
                    hmPd.productUnit,
                    "2019-06-12",
                    0,
                    "2019-06-13"
                )
                productDataItemArr.append(itemStep)
            #----------------2、写入字典用来插入数据库_END----------------#


            
        
        #完结时关闭文件和保存文件
        #----------------生成文件时关闭----------------#
        if isSaveInData != 1:

            nowTime = datetime.datetime.now()
            nowTimeStr = nowTime.strftime("%Y%m%d%H%M%S_s")
            hmPdfSaveName = nowTimeStr +"_"+ hmPdfSaveName+ ".pdf"
            hmPdfSavePath = download_path + hmPdfSaveName

            resultPdfFile = open(hmPdfSavePath,'wb')
            hmPdfWriter.write(resultPdfFile)
            for closeItem in openFileArr :
                closeItem.close()
                os.remove(closeItem.name)
            openFileArr.clear()
            resultPdfFile.close()
        #----------------写入字典用来插入数据库----------------#

        fileOpen.close()
        return hmPdfSaveName

 
if __name__ == '__main__':

    shellValues = sys.argv[1]
    shellArr =  shellValues.split(":")
    isSaveInData = int(shellArr[1])
    fileName = shellArr[0]
    
    #组装数据
    productDataItemArr = []
    autoResult = parse(fileName,productDataItemArr,isSaveInData)
    if isSaveInData == 0:
        print(autoResult)
    
    #写入数据库
    if isSaveInData == 1:
        m = hmDataHandle()
        m.hmInsertBgProduct_sql(hmInsertBgProduct_Str,productDataItemArr)
