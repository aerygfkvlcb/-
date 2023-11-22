import os
import pandas as pd
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'Malgun Gothic'


data1 = pd.ExcelFile('C:/data/실습데이터/KSE_FIN_DATA_2023.xlsx')
data1 = data1.parse(index_col=3)
data_2022 = data1.loc[data1.index == 2022]
data_2021 = data1.loc[(data1.index >= 2021) & (data1.index <= 2022)]


# 수익성
# roe 순이익/자본
data_roe = (data_2022['당기순이익']/data_2022['총자본'])*100
data_roe = data_roe.to_frame(name='roe')
data_roe['score_roe'] = data_roe['roe'].rank(method='min', ascending=True)  # 높을수록 유리, 점수 오름차순
data_2022 = pd.concat([data_2022, data_roe], axis=1)

# roa 순이익/자금
data_roa = (data_2022['당기순이익']/data_2022['총자산'])*100
data_roa = data_roa.to_frame(name='roa')
data_roa['score_roa'] = data_roa['roa'].rank(method='min', ascending=True)  # 높을수록 유리, 점수 오름차순
data_2022 = pd.concat([data_2022, data_roa], axis=1)

# 안정성
# 부채비율 부채/자본
data_debt = (data_2022['총부채']/data_2022['총자본'])*100
data_debt = data_debt.to_frame(name='부채비율')
data_debt['score_부채비율'] = data_debt['부채비율'].rank(method='min', ascending=False)  # 높을수록 불리, 점수 내림차순
data_2022 = pd.concat([data_2022, data_debt], axis=1)

# 자기자본비율 자본/자산
data_capital = (data_2022['총자본']/data_2022['총자산'])*100
data_capital = data_capital.to_frame(name='자기자본비율')
data_capital['score_자기자본비율'] = data_capital['자기자본비율'].rank(method='min', ascending=True)  # 높을수록 유리, 점수 오름차순
data_2022 = pd.concat([data_2022, data_capital], axis=1)

# 성장성
# 매출액증가율 당기매출액-전기매출액 / 전기매출액
data_sales = (data_2021[['Name', '매출액']])
data_sales = (data_sales.iloc[:, 1:] - data_sales.iloc[:, 1:].shift(1)) / data_sales.iloc[:, 1:].shift(1) * 100
data_sales.rename(columns={'매출액': '매출액증가율'}, inplace=True)
data_sales = data_sales.loc[data_sales.index == 2022]
data_sales['score_매출액증가율'] = data_sales['매출액증가율'].rank(method='min', ascending=True)  # 높을수록 유리, 점수 오름차순
data_2022 = pd.concat([data_2022, data_sales], axis=1)

# 매출액증가율 당기매출액-전기매출액 / 전기매출액
data_op = (data_2021[['Name', '영업이익']])
data_op = (data_op.iloc[:, 1:] - data_op.iloc[:, 1:].shift(1)) / data_op.iloc[:, 1:].shift(1) * 100
data_op.rename(columns={'영업이익': '영업이익증가율'}, inplace=True)
data_op = data_op.loc[data_op.index == 2022]
data_op['score_영업이익증가율'] = data_op['영업이익증가율'].rank(method='min', ascending=True)  # 높을수록 유리, 점수 오름차순
data_2022 = pd.concat([data_2022, data_op], axis=1)


data_main = data_2022.loc[:, ['Name', 'roe', 'score_roe', 'roa', 'score_roa', '부채비율', 'score_부채비율',
                              '자기자본비율', 'score_자기자본비율', '매출액증가율', 'score_매출액증가율', '영업이익증가율', 'score_영업이익증가율']]

data_main['총점'] = (data_main['score_roe'] + data_main['score_roa'] + data_main['score_부채비율'] + data_main['score_자기자본비율'] +
                   data_main['score_매출액증가율'] + data_main['score_영업이익증가율']) / 4854 * 100
data_main['총점'] = round(data_main['총점'], 2)

# 1.상위10개 각각 데이터프레임
data_high_all = data_main.nlargest(10, '총점')  # 총점 기준 상위 10개 dataframe
data_high_roe = data_main.nlargest(10, 'roe')  # roe 기준 상위 10개 dataframe
data_high_roa = data_main.nlargest(10, 'roa')  # roa 기준 상위 10개 dataframe
data_high_debt = data_main.nlargest(10, '부채비율')  # 부채비율 기준 상위 10개 dataframe
data_high_capital = data_main.nlargest(10, '자기자본비율')  # 자기자본비율 기준 상위 10개 dataframe
data_high_sales = data_main.nlargest(10, '매출액증가율')  # 매출액증가율 기준 상위 10개 dataframe
data_high_op = data_main.nlargest(10, '영업이익증가율')  # 영업이익증가율 기준 상위 10개 dataframe

# 데이터프레임에서 score열 제거
data_high_all.drop(columns=['score_roe', 'score_roa', 'score_부채비율', 'score_자기자본비율', 'score_매출액증가율', 'score_영업이익증가율'], inplace=True)
data_high_roe.drop(columns=['score_roe', 'score_roa', 'score_부채비율', 'score_자기자본비율', 'score_매출액증가율', 'score_영업이익증가율'], inplace=True)
data_high_roa.drop(columns=['score_roe', 'score_roa', 'score_부채비율', 'score_자기자본비율', 'score_매출액증가율', 'score_영업이익증가율'], inplace=True)
data_high_debt.drop(columns=['score_roe', 'score_roa', 'score_부채비율', 'score_자기자본비율', 'score_매출액증가율', 'score_영업이익증가율'], inplace=True)
data_high_capital.drop(columns=['score_roe', 'score_roa', 'score_부채비율', 'score_자기자본비율', 'score_매출액증가율', 'score_영업이익증가율'], inplace=True)
data_high_sales.drop(columns=['score_roe', 'score_roa', 'score_부채비율', 'score_자기자본비율', 'score_매출액증가율', 'score_영업이익증가율'], inplace=True)
data_high_op.drop(columns=['score_roe', 'score_roa', 'score_부채비율', 'score_자기자본비율', 'score_매출액증가율', 'score_영업이익증가율'], inplace=True)

folder_dir1 = 'C:/data/실습데이터/데이터프레임/'
os.makedirs(folder_dir1, exist_ok=True)

# 2.엑셀로 저장, .xlsx파일 이전의 경로 수정해야함
# data_high_all.to_excel(folder_dir1 + "data_high_all.xlsx")
# data_high_roe.to_excel(folder_dir1 + "data_high_roe.xlsx")
# data_high_roa.to_excel(folder_dir1 + "data_high_roa.xlsx")
# data_high_debt.to_excel(folder_dir1 + "data_high_debt.xlsx")
# data_high_capital.to_excel(folder_dir1 + "data_high_capital.xlsx")
# data_high_sales.to_excel(folder_dir1 + "data_high_sales.xlsx")
# data_high_op.to_excel(folder_dir1 + "data_high_op.xlsx")

data_high_all_name = list(data_high_all['Name'])
data_high_roe_name = list(data_high_roe['Name'])
data_high_roa_name = list(data_high_roa['Name'])
data_high_debt_name = list(data_high_debt['Name'])
data_high_capital_name = list(data_high_capital['Name'])
data_high_sales_name = list(data_high_sales['Name'])
data_high_op_name = list(data_high_op['Name'])


# 최근 1년데이터 추출, 경로 수정 필요
data2 = pd.read_excel('C:/data/실습데이터/kSE수정종가.xlsx', index_col=0, skiprows=range(1, 5640))

folder_dir2 = 'C:/data/실습데이터/주가/'
os.makedirs(folder_dir2, exist_ok=True)
data2.to_excel('C:/data/실습데이터/kSE수정종가_cut.xlsx')

# 1.상위 10개 각각 분류마다 excel로 저장, .xlsx파일 이전의 경로 수정 필요
# data_all_stock = data2.loc[:, data_high_all_name]
# data_all_stock.to_excel(folder_dir2 + "data_all_stock.xlsx")
# data_roe_stock = data2.loc[:, data_high_roe_name]
# data_roe_stock.to_excel(folder_dir2 + "data_roe_stock.xlsx")
# data_roa_stock = data2.loc[:, data_high_roa_name]
# data_roa_stock.to_excel(folder_dir2 + "data_roa_stock.xlsx")
# data_debt_stock = data2.loc[:, data_high_debt_name]
# data_debt_stock.to_excel(folder_dir2 + "data_debt_stock.xlsx")
# data_capital_stock = data2.loc[:, data_high_capital_name]
# data_capital_stock.to_excel(folder_dir2 + "data_capital_stock.xlsx")
# data_sales_stock = data2.loc[:, data_high_sales_name]
# data_sales_stock.to_excel(folder_dir2 + "data_sales_stock.xlsx")
# data_op_stock = data2.loc[:, data_high_op_name]
# data_op_stock.to_excel(folder_dir2 + "data_op_stock.xlsx")


# 2.상위 10개순으로 plt에 그래프 추가
# for i in range(len(data_high_all_name)):
#     data2.plot.line(use_index=True, y=[data_high_all_name[i]])
# for i in range(len(data_high_roe_name)):
#     data2.plot.line(use_index=True, y=[data_high_roe_name[i]])
# for i in range(len(data_high_roa_name)):
#     data2.plot.line(use_index=True, y=[data_high_roa_name[i]])
# for i in range(len(data_high_debt_name)):
#     data2.plot.line(use_index=True, y=[data_high_debt_name[i]])
# for i in range(len(data_high_capital_name)):
#     data2.plot.line(use_index=True, y=[data_high_capital_name[i]])
# for i in range(len(data_high_sales_name)):
#     data2.plot.line(use_index=True, y=[data_high_sales_name[i]])
# for i in range(len(data_high_op_name)):
#     data2.plot.line(use_index=True, y=[data_high_op_name[i]])
# plt.show()
