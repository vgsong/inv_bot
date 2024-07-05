import os
import pandas as pd

mdir = os.path.abspath(os.path.dirname(__name__))
df = pd.read_csv(os.path.join(mdir,'data/wo.csv'))

print(df)