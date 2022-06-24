import json
import pandas
import argparse
import os


def process_attr(attr: str):
    attr = str(attr)
    return attr.split(",")


def get_args():
    parser = argparse.ArgumentParser()

    parser.add_argument("-i", "--input", default="", type=str,
                        help="the input file path to be converted default is src_file")
    parser.add_argument('-o', "--output", default='result', type=str, help="the output dir to save the json file default is result")
    args = parser.parse_args()

    return args


if __name__ == '__main__':
    args = get_args()
    output_file_dir = args.output
    if args.input != "":
        input_file_names = [args.input]
    else:
        # 默认处理文件夹src_file 下的所有excel文件
        input_file_names = [os.path.join('src_file', file) for file in os.listdir('src_file') if file.endswith('xlsx')]
        assert len(input_file_names) > 0, "None input file or the file in src_file dir is specified"
    for input_file_path in input_file_names:
        xl = pandas.ExcelFile(input_file_path)
        sheet_names = xl.sheet_names  # 所有的sheet名称
        for sheet_name in sheet_names:
            with open(os.path.join(output_file_dir, f'{sheet_name}.json'), 'w', encoding='utf8') as fw:
                df = xl.parse(sheet_name)
                df = df.to_dict()
                tmp_result = [{} for attr in df.keys() if isinstance(attr, int)]  # dict 保存中间结果
                obj_index = [i for i in df.keys() if isinstance(i, int)]
                for row_index, parameter in df['parameter'].items():
                    for obj, i in zip(tmp_result, obj_index):
                        obj[parameter] = process_attr(df[i][row_index])  # 实例属性赋值
                json.dump(tmp_result, fw, ensure_ascii=False, indent=4)
