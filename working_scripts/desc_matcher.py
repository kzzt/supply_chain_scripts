from fuzzywuzzy import fuzz
from fuzzywuzzy import process

import pandas as pd


df1 = pd.read_csv(r"c:\users\u6zn\desktop\ref.csv")
df2 = pd.read_csv(r"c:\users\u6zn\desktop\invoice.csv")

df2["key"] = df2.desc.apply(lambda x: [process.extract(x, df1.desc, limit=1)])

df2.merge(df1, left_on="key", right_on="desc")

df2.to_csv(r"c:\users\u6zn\desktop\output_test_file.csv")
