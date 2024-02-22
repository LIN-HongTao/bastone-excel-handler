import tushare as ts

cons = ts.get_apis()
ts.set_token("591e6891f9287935f45fc712bcf62335a81cd6829ce76c21c0fdf7b2")
df = ts.bar("000300", conn=cons, asset="INDEX", start_date="2018-10-01", end_date="2018-12-31")
df = df.sort_index()
import pdb

pdb.set_trace()
