import pandas as pd
from pandas import DataFrame
import requests
import urllib3
import progressbar
import json

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning) # 경고문 예외처리

def get_conf(key):
    with open('conf.json') as f:
        config = json.load(f)

    return config[key]


def get_definition(word):
    key = get_conf('stdict_key')
    url = 'https://stdict.korean.go.kr/api/search.do?'
    method = 'exact' # 일치 검색
    advanced = 'y' # 자세히 찾기 여부
    pos = 1 #명사
    multimedia = [6] # 멀티미디어 포함 x
    definition = []

    params = {
        'key' : key,
        'q' : word,
        'method' : method,
        'advanced' : advanced,
        'pos' : pos,
        'req_type' : 'json',
        'multimedia[]' : multimedia
    }

    res = requests.get(url, params=params, verify=False)
    resp = res.json()
    if resp:
        for i in resp["channel"]["item"]:
            definition.append(i["sense"]["definition"])
            
    
    return definition;


def read_excel(xlse_path, sheetName):
    xls_file = pd.ExcelFile(xlse_path)
    data = xls_file.parse(sheetName)

    return data.to_dict()


def write_excel(filename, df):
    writer = pd.ExcelWriter('./output/' + filename)
    df.to_excel(writer, sheet_name='Sheet1')
    writer.close()


def main():
    excel_dic = read_excel('./input/test.xlsx', 'Sheet1')
    word_dic = excel_dic['word']
    result_list = []
    columns = ['word']
    max_col_len = 0

    bar = progressbar.ProgressBar(maxval= len(word_dic)).start()

    for key in word_dic:
        definition = get_definition(word_dic[key])
        bar.update(key)
        temp_list = []
        temp_list.append(word_dic[key])
        if max_col_len < len(definition):
            max_col_len = len(definition)

        for d in definition :
            temp_list.append(d)
        result_list.append(temp_list)

    for i in range(1, max_col_len + 1) :
        columns.append('definition' + str(i))

    write_excel('result.xlsx', pd.DataFrame(result_list, columns = columns))

    bar.finish()


if __name__ == "__main__":
	main()