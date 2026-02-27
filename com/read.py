import os
import fitz  # PyMuPDF
import pdfplumber
import re
import traceback
import pandas as pd
from datetime import datetime

patterns = {
    '日期': r'(?<=日期[:：]).*(?=回单编号)',
    '回单编号': r'(?<=回单编号[:：]).\d+',
    # '业务类型': r'业务回单\s*（(\w+)）',
    # '付款人户名': r'付款人户名[:：]\s*([^\s:：]{2,})',  # 匹配至少2个非空白字符
    # '付款人账号': r'付款人账号[:：]\s*([0-9]+)',
    '付款人开户行': r'(?<=付款人开户行[:：]).*(?=收款人户名)',
    '收款人户名': r'(?<=收款人户名[:：]).*(?=收款人账号)',
    # '收款人户名': r'收款人户名[:：]\s([^\n]?)\s*(?=收款人账号|$)',
    '收款人账号': r'收款人账号[:：]\s*([0-9]+)',
    '收款人开户行': r'(?<=收款人开户行[:：]).*?(?=币种)',
    # '币种': r'币种[:：]\s*([^\s:：]+)',
    '金额(大写)': r'金额[:：]\s*([壹贰叁肆伍陆柒捌玖拾佰仟万亿圆整]+)',
    '金额(小写)': r'(?<=小写[:：]).*?([0-9,\.]+)',
}


def split_pdf_by_page_fitz(input_pdf_path, output_folder):
    """使用PyMuPDF拆分PDF"""
    os.makedirs(output_folder, exist_ok=True)

    # 打开PDF
    pdf_document = fitz.open(input_pdf_path)

    # 遍历每一页
    for page_num in range(len(pdf_document)):
        # 创建新的PDF文档
        new_pdf = fitz.open()

        # 插入当前页
        new_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)

        # 保存
        output_filename = f"{output_folder}/bank_re_{page_num + 1}.pdf"
        new_pdf.save(output_filename)
        new_pdf.close()

        print(f"已保存: {output_filename}")

    pdf_document.close()
    print("拆分完成！")


def read_chinese_pdf(files):
    """
    专门读取中文PDF，优化中文文本提取
    """
    print("读取中文PDF")
    result_list = []
    for index, pdf_file in enumerate(files):
        try:
            with pdfplumber.open(pdf_file) as pdf:
                # all_text = ""
                for page_num, page in enumerate(pdf.pages, 1):
                    # print(f"\n处理第 {page_num} 页...")

                    # 尝试不同的提取策略
                    strategies = [
                        ("常规", lambda p: p.extract_text()),
                        ("Layout", lambda p: p.extract_text(layout=True)),
                        ("使用空格", lambda p: p.extract_text(extra_attrs=["fontname", "size"])),
                    ]

                    best_text = ""
                    best_strategy = ""
                    for strategy_name, extract_func in strategies:
                        try:
                            text = extract_func(page)
                            if text and len(text.strip()) > len(best_text):
                                # 检查是否包含中文
                                chinese_chars = re.findall(r'[\u4e00-\u9fff]+', text)
                                if chinese_chars:
                                    best_text = text
                                    best_strategy = strategy_name
                        except:
                            continue

                    if best_text:
                        # print(f"使用策略: {best_strategy}")
                        # 清理文本
                        cleaned_text = clean_chinese_text(best_text)
                        # 查找发票信息
                        find_info = find_invoice_info(cleaned_text, index + 1, patterns)
                        find_info['pdf_路径'] = pdf_file
                        # all_text += f"=== 第 {page_num} 页 ===\n{cleaned_text}\n\n"
                        # print(all_text)
                    else:
                        print("未提取到有效文本")
                result_list.append(find_info)
        except Exception as e:
            print(f"读取中文PDF出错: {traceback.format_exc()}")
            continue
    return result_list


def clean_chinese_text(text):
    """清理中文文本"""
    # 替换多个空格
    text = re.sub(r'\s+', ' ', text).replace(' ', '')
    # 修复常见的中文标点问题
    text = re.sub(r'([\u4e00-\u9fff])\s+([\u4e00-\u9fff])', r'\1\2', text)
    # 修复数字和中文之间的空格
    text = re.sub(r'(\d)\s+([\u4e00-\u9fff])', r'\1\2', text)
    text = re.sub(r'([\u4e00-\u9fff])\s+(\d)', r'\1\2', text)
    return text.strip()


# 查找信息
def find_invoice_info(text, page_num, patterns):
    """查找发票信息"""

    found_info = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            # print(match)
            try:
                found_info[key] = match.group()
            except Exception as e:
                print(traceback.format_exc())
                print(f"ERRO当前key{key},在票据中找不到")
    if found_info:
        # print(f"  第 {page_num} 页发现的发票信息:")
        # for key, value in found_info.items():
        #     # print(f"    {key}: {value}")

        return found_info
    else:
        return None


# 读取交易信息
def read_transaction_info_for_excel(excel_file):
    data_df = pd.read_excel(excel_file, dtype=str)
    filter_df = data_df.loc[~data_df['对方户名'].isna() & ~data_df['对方开户行'].isna()]
    return filter_df


def get_files_length(file_path):
    all_items = []
    for basename in os.listdir(file_path):
        all_items.append(os.path.join(file_path, basename))
    return all_items


# 新比对逻辑
def compare_all_data(pdf_list, transaction_df, pz_df):
    """
    统一对比PDF数据、交易明细表和凭证表
    只有三个数据源都能相互匹配才算成功
    """
    # 预处理数据
    transaction_df = transaction_df.copy()
    pz_df = pz_df.copy()

    # 转换数据类型
    transaction_df['付款金额'] = pd.to_numeric(transaction_df['付款金额'], errors='coerce')
    pz_df['付款金额'] = pd.to_numeric(pz_df['付款金额'], errors='coerce')

    # 转换日期格式
    if '交易日期' in transaction_df.columns:
        transaction_df['交易日期'] = pd.to_datetime(transaction_df['交易日期']).dt.date

    if '时间' in pz_df.columns:
        pz_df['交易日期'] = pd.to_datetime(pz_df['时间']).dt.date

    # 为每个数据源添加唯一标识
    pdf_records = []
    for i, item in enumerate(pdf_list):
        item['数据源'] = 'PDF'
        item['唯一标识'] = f'PDF_{i}'
        # 清理金额字段
        if '金额(小写)' in item:
            item['金额'] = float(str(item['金额(小写)']).replace(',', ''))
        pdf_records.append(item)

    transaction_df['数据源'] = '交易明细表'
    transaction_df['唯一标识'] = ['TRANS_' + str(i) for i in range(len(transaction_df))]

    pz_df['数据源'] = '凭证表'
    pz_df['唯一标识'] = ['PZ_' + str(i) for i in range(len(pz_df))]

    # 匹配结果存储
    success_list = []  # 存储三者都匹配成功的数据
    failure_list = []  # 存储匹配失败的数据

    # 1. 创建匹配字典，记录每个PDF记录匹配到的交易明细和凭证
    pdf_matches = {}

    for pdf_item in pdf_records:
        pdf_key = f"{pdf_item.get('收款人户名', '')}_{pdf_item.get('金额', 0)}_{pdf_item.get('日期', '')}"
        pdf_matches[pdf_key] = {
            'pdf_record': pdf_item,
            'trans_match': None,
            'pz_match': None,
            'matched': False
        }

        # 尝试匹配交易明细表
        trans_matches = []
        for _, trans_row in transaction_df.iterrows():
            # 金额匹配（允许微小误差）
            amount_match = abs(float(trans_row['付款金额']) - pdf_item['金额']) <= 0.001

            # 对方户名匹配
            name_match = trans_row.get('对方户名', '') == pdf_item.get('收款人户名', '')

            # 日期匹配
            trans_date = trans_row.get('交易日期')
            pdf_date = pdf_item.get('日期')
            if isinstance(pdf_date, str):
                pdf_date = datetime.strptime(pdf_date, "%Y-%m-%d").date() if pdf_date else None

            date_match = str(trans_date) == str(pdf_date)

            if amount_match and name_match and date_match:
                trans_matches.append(trans_row)

        # 尝试匹配凭证表
        pz_matches = []
        for _, pz_row in pz_df.iterrows():
            # 金额匹配
            amount_match = abs(float(pz_row['付款金额']) - pdf_item['金额']) <= 0.001

            # 日期匹配
            pz_date = pz_row.get('交易日期')
            pdf_date = pdf_item.get('日期')
            if isinstance(pdf_date, str):
                pdf_date = datetime.strptime(pdf_date, "%Y-%m-%d").date() if pdf_date else None

            date_match = str(pz_date) == str(pdf_date)

            # 摘要中包含对方户名
            summary_match = False
            if '摘要' in pz_row and pd.notna(pz_row['摘要']):
                account_name = pdf_item.get('收款人户名', '')
                if account_name and account_name in str(pz_row['摘要']):
                    summary_match = True

            if amount_match and date_match and summary_match:
                pz_matches.append(pz_row)

        # 记录匹配结果
        if trans_matches and pz_matches:
            # 三者匹配成功
            pdf_matches[pdf_key]['trans_match'] = trans_matches[0]
            pdf_matches[pdf_key]['pz_match'] = pz_matches[0]
            pdf_matches[pdf_key]['matched'] = True

            # 构建成功记录
            success_record = {
                '匹配状态': '三者匹配成功',
                'PDF数据': pdf_item,
                '交易明细数据': trans_matches[0].to_dict(),
                '凭证表数据': pz_matches[0].to_dict(),
                '匹配金额': pdf_item['金额'],
                '匹配日期': pdf_item.get('日期'),
                '对方户名': pdf_item.get('收款人户名', '')
            }
            success_list.append(success_record)

            # 标记已匹配的记录
            transaction_df = transaction_df[transaction_df['唯一标识'] != trans_matches[0]['唯一标识']]
            pz_df = pz_df[pz_df['唯一标识'] != pz_matches[0]['唯一标识']]
        else:
            # 匹配失败，记录失败原因
            failure_reason = []
            if not trans_matches:
                failure_reason.append('缺少交易明细表匹配')
            if not pz_matches:
                failure_reason.append('缺少凭证表匹配')

            failure_record = {
                '数据来源': 'PDF数据',
                '匹配状态': '部分匹配失败',
                '失败原因': '；'.join(failure_reason),
                'PDF数据': pdf_item,
                '匹配到的交易明细数量': len(trans_matches),
                '匹配到的凭证表数量': len(pz_matches)
            }
            failure_list.append(failure_record)

    # 2. 检查剩余的未匹配的交易明细记录
    for _, trans_row in transaction_df.iterrows():
        # 检查是否有匹配的凭证表
        pz_matches = []
        for _, pz_row in pz_df.iterrows():
            # 金额匹配
            amount_match = abs(float(pz_row['付款金额']) - float(trans_row['付款金额'])) <= 0.001

            # 日期匹配
            pz_date = pz_row.get('交易日期')
            trans_date = trans_row.get('交易日期')
            date_match = str(pz_date) == str(trans_date)

            # 摘要中包含对方户名
            summary_match = False
            if '摘要' in pz_row and pd.notna(pz_row['摘要']):
                account_name = trans_row.get('对方户名', '')
                if account_name and account_name in str(pz_row['摘要']):
                    summary_match = True

            if amount_match and date_match and summary_match:
                pz_matches.append(pz_row)

        # 记录匹配结果
        if pz_matches:
            # 交易明细表和凭证表匹配，但缺少PDF
            failure_record = {
                '数据来源': '交易明细表+凭证表',
                '匹配状态': '缺少PDF匹配',
                '失败原因': '交易明细表和凭证表匹配成功，但缺少对应的PDF数据',
                '交易明细数据': trans_row.to_dict(),
                '凭证表数据': pz_matches[0].to_dict(),
                '匹配金额': trans_row['付款金额'],
                '匹配日期': trans_row.get('交易日期'),
                '对方户名': trans_row.get('对方户名', '')
            }
            failure_list.append(failure_record)

            # 标记已匹配的凭证表记录
            pz_df = pz_df[pz_df['唯一标识'] != pz_matches[0]['唯一标识']]
        else:
            # 只有交易明细表，缺少PDF和凭证表
            failure_record = {
                '数据来源': '交易明细表',
                '匹配状态': '缺少PDF和凭证表匹配',
                '失败原因': '只有交易明细表数据，缺少对应的PDF和凭证表数据',
                '交易明细数据': trans_row.to_dict()
            }
            failure_list.append(failure_record)

    # 3. 检查剩余的未匹配的凭证表记录
    for _, pz_row in pz_df.iterrows():
        failure_record = {
            '数据来源': '凭证表',
            '匹配状态': '缺少PDF和交易明细表匹配',
            '失败原因': '只有凭证表数据，缺少对应的PDF和交易明细表数据',
            '凭证表数据': pz_row.to_dict()
        }
        failure_list.append(failure_record)

    # 4. 输出结果到Excel
    write_success_to_excel(success_list, '三者匹配成功.xlsx')
    write_failure_to_excel(failure_list, '比对未成功.xlsx')

    print(f"匹配完成：")
    print(f"- 三者匹配成功：{len(success_list)} 条")
    print(f"- 匹配失败：{len(failure_list)} 条")

    return success_list, failure_list


# 写入比对成功的
def write_success_to_excel(success_list, filename):
    """将成功匹配的数据写入Excel"""
    if success_list:
        # 扁平化成功数据
        flattened_data = []
        for record in success_list:
            flat_record = {
                '匹配状态': record['匹配状态'],
                '匹配金额': record['匹配金额'],
                '匹配日期': record['匹配日期'],
                '对方户名': record['对方户名'],
                'PDF_收款人户名': record['PDF数据'].get('收款人户名', ''),
                'PDF_金额': record['PDF数据'].get('金额', ''),
                'PDF_日期': record['PDF数据'].get('日期', ''),
                'PDF_路径': record['PDF数据'].get('pdf_路径', ''),
                '交易明细_对方户名': record['交易明细数据'].get('对方户名', ''),
                '交易明细_金额': record['交易明细数据'].get('付款金额', ''),
                '交易明细_日期': record['交易明细数据'].get('交易日期', ''),
                '凭证表_摘要': record['凭证表数据'].get('摘要', ''),
                '凭证表_金额': record['凭证表数据'].get('付款金额', ''),
                '凭证表_日期': record['凭证表数据'].get('时间', '')
            }
            flattened_data.append(flat_record)

        df = pd.DataFrame(flattened_data)
        df.to_excel(filename, index=False)
        print(f"成功数据已保存到：{filename}")


# 写入比对失败的
def write_failure_to_excel(failure_list, filename):
    """将匹配失败的数据写入Excel"""
    if failure_list:
        df = pd.DataFrame(failure_list)
        df.to_excel(filename, index=False)
        print(f"失败数据已保存到：{filename}")


def main():
    dir_path = os.path.dirname(os.path.dirname(__file__))
    print(dir_path)
    input_pdf_path = os.path.join(dir_path, 'files', '银行回单.pdf')
    # 交易明细表
    transaction_excel_path = os.path.join(dir_path, 'files', '工行交易明细.xlsx')
    # 凭证表
    pz_excel_path = os.path.join(dir_path, 'files', '凭证表.xlsx')
    # 拆分pdf
    try:
        split_file = os.path.join(dir_path, 'com', '拆分结果')
        files = get_files_length(split_file)
    except Exception as e:
        print('未创建拆分结果文件夹')
        split_pdf_by_page_fitz(input_pdf_path, "拆分结果")
        split_file = os.path.join(dir_path, 'com', '拆分结果')
        files = get_files_length(split_file)
    # 拆分后的pdf结果列表
    pdf_list = read_chinese_pdf(files)
    # 读取凭证表
    pz_df = pd.read_excel(pz_excel_path, dtype={'付款金额': float, '时间': str}, skiprows=1)
    print(pz_df.head())
    print(input_pdf_path)
    # print(pdf_list)
    # 核对对方户名，对方开户行，付款金额，交易时间
    # 读取交易明细
    transaction_df = read_transaction_info_for_excel(transaction_excel_path)
    compare_all_data(pdf_list, transaction_df, pz_df)
    print('===============================比对完成========================================')


if __name__ == '__main__':
    main()
    # print(list(patterns.keys()))
