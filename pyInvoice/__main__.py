import argparse
import logging
from datetime import datetime
import os   
from logging import StreamHandler
import pandas as pd
from typing import Tuple, Dict, Any
from pyInvoice import constant
from pyInvoice.template import Template
from pyInvoice.data_manager import *
from tqdm import tqdm

#ロガー
def getLogger(base_dir) -> logging.Logger:
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter('[%(asctime)s] %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    console_handler = StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # ファイルハンドラ作成
    file_handler = logging.FileHandler(f"{constant.attr_output_dir}/{base_dir}/{constant.attr_log_file}", mode='a', encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    
	# ロガーに追加
    logger.addHandler(file_handler)
    return logger

def import_setting(file_path: str) -> Tuple[bool, Dict[str, Any] | None]:
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"指定されたファイルが存在しません: {file_path}")

        logger.info(f"設定ファイル読み込み開始: {file_path}")
        # すべてのシートを一度に読み込む
        sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl', dtype=str)

        setting_params = { }

        # 'ファイル'シートからTableSourceの辞書を作成
        if not "ファイル" in sheets:
            raise ValueError("設定ファイルに「ファイル」シートが存在しません")
        
        file_sheet = sheets['ファイル']
        ret = {
            row['ファイル名称']: TableSource(path=row['ファイルパス'], sheet_name=row['シート名称'])
            for _, row in file_sheet.iterrows()
        }
        setting_params[constant.attr_str_file_params_key] = ret

        # '設定'シートから設定項目を追加
        if not "設定" in sheets:
            raise ValueError("設定ファイルに「設定」シートが存在しません")
        setting_sheet = sheets['設定']
        ret = {
            row['項目']: row['設定']
            for _, row in setting_sheet.iterrows()
        }
        setting_params[constant.attr_str_setting_params_key] = ret

        #テンプレート情報を追加
        if not "テンプレート" in sheets:
            raise ValueError("設定ファイルに「テンプレート」シートが存在しません")
        setting_sheet = sheets['テンプレート']
        ret = {
            row['名称']: {'識別記号' : row['識別記号'], 'テンプレートファイル' : row['テンプレートファイル'], }
            for _, row in setting_sheet.iterrows()
        }
        setting_params[constant.attr_str_temp_params_key] = ret
        
        logger.info(f"設定ファイル読み込み完了")
        return True, setting_params
    except Exception as e:
        logger.error("設定ファイル読み込み失敗", exc_info=False)
        print(e)
        return False, None

def import_packing_list(file_path : str, sheet_name : str) -> Tuple[bool, PackingList | None]:
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"指定されたファイルが存在しません: {file_path}")

        logger.info(f"パッキング読み込み開始: {file_path}")
        sheet = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', dtype=str, header=None).fillna("")
        date, nouhin_no = sheet.iloc[5, 3], sheet.iloc[6, 3]

        packing_list = PackingList(date=date, nouhin_no=nouhin_no)

        jan, row = "jan", 13 - 1
        while(jan and not pd.isna(jan)):
            jan = sheet.iloc[row, 1]
            amount = sheet.iloc[row, 3]
            no = sheet.iloc[row, 4]
            if(not pd.isna(jan) and jan != ""):
                packing = Packing(jan=jan, amount=amount, no=no)
                packing_list.add(packing=packing)
            row += 1
        logger.info(f"{len(packing_list.packings)}件の設定ファイル読み込み完了")
        return True, packing_list
    except ValueError as ve:
        logger.error(f"パッキングファイル読み込み失敗: シート{sheet_name}が存在しません", exc_info=False)
        return False, None
    except Exception as e:
        logger.error("パッキングファイル読み込み失敗", exc_info=False)
        return False, None
    
def import_item_master(file_path : str, sheet_name : str) -> Tuple[bool, bool | None]:

    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"指定されたファイルが存在しません: {file_path}")

        logger.info(f"商品マスタ読み込み開始: {file_path}")
        sheet = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', dtype=str, header=2).fillna("")
        
        item_master = ItemMaster()
        for _, row in tqdm(sheet.iterrows(), desc="【商品マスタ読込】", total=len(sheet)):            
            item = Item(row=row)
            item_master.add(item)
        logger.info(f"{len(item_master.items)}件の商品マスタファイル読み込み完了")
        return True, item_master
    except ValueError as ve:
        logger.error(f"商品マスタファイル読み込み失敗: シート{sheet_name}が存在しません", exc_info=False)
        return False, None
    except Exception as e:
        logger.error("商品マスタファイル読み込み失敗", exc_info=False)
        return False, None
    
def import_data(options) -> bool:

    file_path = constant.attr_str_setting_file_path
    success, setting_dict = import_setting(file_path)
    if(not success):
        return False, None, None, None
    
    src : TableSource = setting_dict.get(constant.attr_str_file_params_key, {}).get(constant.attr_str_packing_key)    
    success, packing_data = import_packing_list(src.path, src.sheet_name)
    if(not success):
        return False, None, None, None
    
    src : TableSource = setting_dict.get(constant.attr_str_file_params_key, {}).get(constant.attr_str_item_master_key)
    success, item_master = import_item_master(src.path, src.sheet_name)
    if(not success):
        return False, None, None, None
    
    return True, setting_dict, packing_data, item_master

def generate(base_dir, setting_dict, packing_data, item_master):
    customer = setting_dict.get(constant.attr_str_setting_params_key, {}).get(constant.attr_str_customer_key)    
    temp_setting = setting_dict.get(constant.attr_str_temp_params_key, {}).get(customer)

    if(temp_setting):
        temp = Template(
            name=customer,
            label=temp_setting.get(constant.attr_str_label_key),
            base_file_path=temp_setting.get(constant.attr_str_temp_file_path_key),
            output_dir=f"{constant.attr_output_dir}/{base_dir}", 
            setting_dict=setting_dict.get(constant.attr_str_setting_params_key, {})
        )
    else:
        temp = None
        logger.warning(f"無効な宛先です: {customer}")

    # if(customer == "ZOZO"):
    #     temp = Template(
    #         name = "ZOZOTOWN", 
    #         label = "zozo",
    #         sheet_name = "ZOZO",
    #         base_file_path="./template/ZOZOTOWN.xlsx",
    #         output_dir=f"{constant.attr_output_dir}/{base_dir}", 
    #         setting_dict=setting_dict
    #     )
    # elif(customer == "MEGASEEK"):
    #     temp = Template(
    #         name = "MEGASEEK", 
    #         label = "dms",
    #         sheet_name = "MEGASEEK",
    #         base_file_path="./template/MEGASEEK.xlsx", 
    #         output_dir=f"{constant.attr_output_dir}/{base_dir}", 
    #         setting_dict=setting_dict
    #     )
    # elif(customer == "RAKUTEN_FASION"):
    #     temp = Template(
    #         name = "RAKUTEN_FASION", 
    #         label = "rf",
    #         sheet_name = "RAKUTEN_FASION",
    #         base_file_path="./template/RAKUTEN_FASION.xlsx", 
    #         output_dir=f"{constant.attr_output_dir}/{base_dir}", 
    #         setting_dict=setting_dict
    #     )
    # elif(customer == "RSL"):
    #     temp = Template(
    #         name = "RSL", 
    #         label = "rsl",
    #         sheet_name = "RSL",
    #         base_file_path="./template/RSL.xlsx", 
    #         output_dir=f"{constant.attr_output_dir}/{base_dir}", 
    #         setting_dict=setting_dict
    #     )
    # else:
    #    temp = None
    #    logger.warning(f"無効な宛先です{customer}")

    if(temp):
        temp.write(packing_data, item_master, logger)

    return 

def preprocess(base_dir):
 
    os.makedirs(f"{constant.attr_output_dir}/{base_dir}", exist_ok=True)

    return 

if __name__ == "__main__":

    base_dir = datetime.now().strftime("%Y%m%d_%H%M%S")
    preprocess(base_dir)

    logger = getLogger(base_dir)
    parser = argparse.ArgumentParser()

    parser.add_argument("-d", "--debug", action="store_true", help="debug flag")
    # parser.add_argument("-i", "--input", required=True, help="入力ファイルのパスを指定してください")
    # parser.add_argument("-o", "--output", required=False, help="出力ファイルのパスを指定してください")

    options = parser.parse_args()
    logger.info("========== import_data 開始 ==========")
    success, setting_dict, packing_data, item_master = import_data(options)
    logger.info("========== import_data 終了 ==========\n")

    if(success):
        logger.info("========== generate 開始 ==========")
        generate(base_dir, setting_dict, packing_data, item_master)
        logger.info("========== generate 終了 ==========")
    else:
        logger.error("ファイル読込エラー")
    
    # print("press enter...")
    # input()
    