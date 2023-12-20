import os
import pandas as pd # pip install pandas
# import glob
from datetime import date
# from openpyxl import load_workbook
from assets.settings import excel_password

import msoffcrypto
import io
smartstore_encrypted = io.BytesIO()
with open('./assets/testdata/smartstore_download.xlsx', 'rb') as f:
    excel = msoffcrypto.OfficeFile(f)
    excel.load_key(excel_password)
    excel.decrypt(smartstore_encrypted)

today = date.today()

""" 
CODEFLOW
#1 Download excel file from Naver Smartstore
#2 Download excel file from Coupang Wing
#3 Produce upload excel file for Lois Parcel
#4 Download tracking data from Lois Parcel
#5 Upload tracking data to Naver Smartstore
#6 Upload tracking data to Coupang Wing
"""


""" FILE DIALOG """
# import tkinter as tk
# from tkinter import filedialog, PhotoImage
# root = tk.Tk()
# root.wm_title('Eskins LoisParcel Excel Maker')
# img = PhotoImage(file='./assets/eskins_logo.ico')
# root.iconphoto(False, img)
# root.withdraw()
# file_path = filedialog.askopenfilename()

# loading templates
loisparcel_upload_df = pd.read_excel("./assets/template/loisparcel upload/loisparcel_upload_template.xlsx")
smartstore_upload_df = pd.read_excel("./assets/template/smartstore upload/smartstore_upload_template.xls")
cwing_upload_df = pd.read_excel("./assets/template/cwing upload/cwing_upload_template.xlsx")

# loading data
cwing_download_df = pd.read_excel("./assets/testdata/cwing_download.xlsx")
smartstore_download_df = pd.read_excel(smartstore_encrypted, skiprows=[0])
loisparcel_download_df = pd.read_excel("./assets/testdata/loisparcel_download.xlsx")
# print(smartstore_download_df.loc[[1]])
"""
#3 LOISPARCEL UPLOAD
INPUT: cwing_download_df, smartstore_download_df
OUTPUT: loisparcel_upload_df
"""

loisparcel_upload_df[[
    '고객주문번호',
    '받는분성명',
    '품목명',
    '내품명',
    '내품수량',
    '배송메세지1',
    '받는분전화번호',
    '우편번호',
    '받는분주소(전체, 분할)'
]] = cwing_download_df[[
    '주문번호',
	'수취인이름',
	'등록상품명',
	'등록옵션명',
	'구매수(수량)',
	'배송메세지',
	'수취인전화번호',
	'우편번호',
	'수취인 주소'
]]

smartstore_upload_df[[
    '택배사',	#1
    '상품주문번호', #2
    ]] = smartstore_download_df[[
		'택배사',	#1
		'주문번호', #2
        ]]
    
smartstore_upload_df[[
    '배송방법',	#1
    ]] = smartstore_download_df[[
		'배송방법',	#1
        ]]

smartstore_download_df.rename(columns={
    '주문번호': '고객주문번호',
    '수취인명': '받는분성명',
    '상품명': '품목명',
    '옵션정보': '내품명',
    '수량': '내품수량',
    '배송메세지': '배송메세지1',
    '수취인연락처1': '받는분전화번호',
    '우편번호': '우편번호',
    '기본배송지': '받는분주소(전체, 분할)',
    '상세배송지': '받는분상세주소(분할)' # only for smartstore
}, inplace=True)

loisparcel_upload_df = pd.concat([loisparcel_upload_df, smartstore_download_df])
loisparcel_upload_df.reset_index(inplace=True, drop=True)

"""
#5 COUPANG WING TRACKING UPLOAD
INPUT: loisparcel_download_df
OUTPUT: cwing_upload_df
"""
cwing_upload_df[[
    '번호',	#1
    '묶음배송번호', #2
    '주문번호', #3
    '택배사', #4
    '분리배송 Y/N', #5
    '수취인이름', #6
    '주문일', #7
    '등록상품명', #8
    '등록옵션명', #9
    '옵션ID', #10
    '구매수(수량)', #11
    '수취인전화번호', #12
    '우편번호', #13
    '수취인 주소', #14
    '배송메세지', #15
	'운송장번호' #16
    ]] = cwing_download_df[[
		'번호',	#1
		'묶음배송번호', #2
		'주문번호', #3
		'택배사', #4
		'분리배송 Y/N', #5
		'수취인이름', #6
		'주문일', #7
		'등록상품명', #8
		'등록옵션명', #9
		'옵션ID', #10
		'구매수(수량)', #11
		'수취인전화번호', #12
		'우편번호', #13
		'수취인 주소', #14
		'배송메세지', #15
		'운송장번호' #16
        ]]
    
""" 
#6 smartstore TRACKING UPLOAD
INPUT: loisparcel_download_df
OUTPUT: smartstore_upload_df
"""
# Pre-Processing
# print(smartstore_upload_df)

smartstore_upload_df[[
    '송장번호', #1
    ]] = loisparcel_download_df[[
		'운송장번호', #1
        ]]


"""
OUTPUT (UPLOAD DATA)
1. Lois Parcel shipment upload
2. cwing Wing tracking upload
3. smartstore Smartstore tracking upload
"""
loisparcel_upload_df.to_excel('loisparcel upload/cj_upload' + str(today) + '.xlsx', index=False)

cwing_upload_df.to_excel('cwing upload/cwing_tracking' + str(today) + '.xlsx', index=False)
smartstore_upload_df.to_excel('smartstore upload/smartstore_tracking' + str(today) + '.xlsx', index=False)

""" 
EXCEL FORMAT
#1
#2
#3
#4 loisparcel DOWNLOAD: No	선택	접수순서	예약구분	상태	집화예정일자	운송장번호	집화예정점소	보내는분	보내는분전화번호	운임구분	박스타입"	수량	내품수량	기본운임	기타운임	운임합계	고객주문번호	배송계획점소	받는분	받는분전화번호	받는분우편번호	받는분주소	상품코드	상품명	단품코드	단품명	배송메시지
#5 smartstore UPLOAD:	상품주문번호	배송방법	택배사	송장번호
#6 cwing UPLOAD:
"""