from dateutil.relativedelta import relativedelta
from datetime import datetime, timedelta
from openpyxl.utils.datetime import from_excel

class converterClass:
    def __init__(self) -> None:
        self.encoder = dict()
        self.decoder = dict()
        pass
    def build(self, start_time, n):
        for i in range(n):
            time = start_time
            datetime1 = datetime.combine(datetime.today(), start_time) + timedelta(minutes=30*i)
            time = datetime1.time()
            self.encoder[time] = i
            self.decoder[i] = time
        return 

def to_int(value):
    try :
        return int(float(value))
    except:
        return

def to_month(value : datetime) -> str:
    return f"{value.year}_{value.month}_1"

def is_numeric(value) -> bool:
    try:
        float(value)
        return True
    except:
        return False

def to_datetime(value : str) -> datetime:
    for format in ['%Y/%m/%d %H:%M', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S']:
        try:
            date = datetime.strptime(value, format)
            return date, True
        except:
            continue
    return None, False

def to_datetime_from_excel(value):
    if(is_numeric(value)):
        value = float(value)

    if isinstance(value, (int, float)):
        date_value = from_excel(value)  # datetime型に変換
    elif isinstance(value, str):
        try:
            date_value = datetime.strptime(value, '%Y-%m-%d')
        except ValueError:
            date_value = datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
    elif isinstance(value, datetime):
        date_value = value
    else:
        raise ValueError(f"日付の形式が不正です: {value}")
    return date_value

def judge(str) : 
    return "○" if str else "×"
