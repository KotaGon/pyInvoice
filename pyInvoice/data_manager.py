from collections import defaultdict
from pyInvoice import utils

class TableSource:
    def __init__(self, path="", sheet_name=""):
        self.path = path
        self.sheet_name = sheet_name

class Packing:
    def __init__(self, jan : str, no : int, amount : int):
        self.jan = jan
        self.no = no
        self.amount = utils.to_int(amount)

class PackingList:
    def __init__(self, date, nouhin_no):
        self.date = date
        self.nouhin_no = nouhin_no
        self.packings = []
        self.packing_set = defaultdict(list)  # 自動でリストを初期化してくれる

    def add(self, packing: Packing) -> None:
        self.packings.append(packing)
        self.packing_set[packing.no].append(packing)
    
class Item:
    # JANCD	品名	備考	ブランド名	ブランド品番/ASIN/SKU	カラー	サイズ		定価（税込）	定価（税抜）	仕入(円)	仕入(元)	仕入(USD)	入庫合計	出庫合計	現在庫
    def __init__(self, row={}):
        self.jan        = row.get("JANCD", "")
        self.name       = row.get("品名", "")
        # 備考はスキップ
        self.brand      = row.get("ブランド名", "")
        self.brand_code = row.get("ブランド品番/ASIN/SKU", "")
        self.color      = row.get("カラー", "")
        self.size       = row.get("サイズ", "")
        self.price_incl_tax = utils.to_int(row.get("定価（税込）", 0))
        self.price          = utils.to_int(row.get("定価（税抜）", 0))
        # self.cost_jpy  = utils.to_int(row.get("仕入(円)", 0))
        # self.stock     = utils.to_int(row.get("現在庫", 0))

        self.row = row  # 元のデータ保持


class ItemMaster:
    def __init__(self):
        self.items = list()
        self.item_dict = dict()
    def add(self, item : Item) -> None:
        self.items.append(item)
        self.item_dict[item.jan] = item

class ItemConv:
    # マスターコード	ブランド名	商品名	カラー	サイズ	税抜き上代
    def __init__(self, row={}):
        self.row = row  # 元のデータを保持

        self.master_code = row.get("マスターコード", "")
        self.brand       = row.get("ブランド名", "")
        self.name        = row.get("商品名", "")
        self.color       = row.get("カラー", "")
        self.size        = row.get("サイズ", "")
        self.price       = utils.to_int(row.get("税抜き上代", 0))

class ItemConvMaster:
    def __init__(self):
        self.items = list()
        self.item_dict = dict()
    def add(self, item : ItemConv) -> None:
        self.items.append(item)
        self.item_dict[item.master_code] = item
