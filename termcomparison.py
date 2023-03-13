import os
import re
from pathlib import Path

import pandas as pd
from urllib.request import urlopen
from urllib.request import urlretrieve
from urllib.parse import quote


link_profix = "https://scicrunch.org/scicrunch/interlex/search?q="

dataFrame = pd.DataFrame()
human_nerves_df = pd.read_excel("human-nerves_with-models.xlsx")
FC_df = pd.read_excel("FCannotationsv2.xlsx", sheet_name="Neural")
human_nerves_df.columns = ["ID", "TYPE", "LABEL", "Model", "Notes"]
human_nerves_df["ScicrunchLabel"] = ''
human_nerves_df["inFC"] = ''


def get_html(url):
    try:
        page = urlopen(url)
        htmlcode = page.read()
        return htmlcode.decode('utf-8')
    except:
        print("Get_html Url Error", url)


def get_title(htmlcode, code):
    titleReg = f'searchTerm={code}">(\s*.*)<\/a>'
    title = getRegResult(htmlcode, titleReg)
    if title:
        return title[0]


def get_prefer_ID(htmlcode):
    regx = f'<p><b>Preferred ID:<\/b> (\s*.*)&nbsp;&nbsp;&nbsp;&nbsp;'
    prefer_ID = getRegResult(htmlcode, regx)
    if prefer_ID:
        return prefer_ID[0]


def getRegResult(htmlcode, reg):
    reg = re.compile(reg)
    results = reg.findall(htmlcode)
    return list(results)


def addScicrunchLabel(df, model, title):
    index = df.index[df['Model'] == model].tolist()
    for i in index:
        df.at[i, "ScicrunchLabel"] = title
    return df


def addInFC(df, model, title):
    index = df.index[df['Model'] == model].tolist()
    for i in index:
        df.at[i, "inFC"] = title
    return df


def get_index_inFC(id, prefer_ID):
    index_list = FC_df[(FC_df["Models"] == id) | (FC_df["Models"] == prefer_ID)].index.values.tolist()
    new_list = [str(x+2) for x in index_list]
    return ','.join(new_list)
# print(FC_df)


# human_nerves_df = human_nerves_df.head(3)
human_nerves_df.ScicrunchLabel.fillna('', inplace=True)
for i in human_nerves_df["Model"]:
    if isinstance(i, str) and i.strip() != "Term requested":
        link = link_profix + quote(i)
        print(link)
        htmlcode = get_html(link)
        title = get_title(htmlcode, i)
        prefer_ID = get_prefer_ID(htmlcode)
        print(title)
        print(i, prefer_ID)
        inFCIndex = get_index_inFC(i, prefer_ID)
        human_nerves_df = addScicrunchLabel(human_nerves_df, i, title)
        human_nerves_df = addInFC(human_nerves_df, i, inFCIndex)


def change_backcolor(x):
    color_df = pd.Series('', index=x.index)
    scicrunchLabel_bgc = f'background-color: {"#99FF99" if x["LABEL"].strip().lower() == x["ScicrunchLabel"].strip().lower() else "#FF7C80" }'
    color_df['ScicrunchLabel'] = scicrunchLabel_bgc
    model_bgc = f'{"background-color: #FF7C80" if not isinstance(x["Model"], str) or x["Model"].strip() == "Term requested" else "" }'
    color_df['Model'] = model_bgc
    return color_df


human_nerves_df.style.apply(change_backcolor, axis=1).to_excel("human_nerves_output.xlsx")
