import pandas as pd
from docxtpl import DocxTemplate
import json
order_df = json.load(open("json/order_info.json"))
print(len(order_df))