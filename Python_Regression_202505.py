# -*- coding: utf-8 -*-
import math
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import statsmodels.api as sm
import statsmodels.stats.api as sms
import statsmodels.robust.norms as robust_norms
from statsmodels.iolib.summary2 import summary_col
from statsmodels.stats.diagnostic import linear_rainbow, het_breuschpagan, het_goldfeldquandt
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.outliers_influence import variance_inflation_factor
import scipy.stats as stats
from scipy.stats import shapiro, jarque_bera, skew, kurtosis
import sys
from tkinter import filedialog
import re
from pathlib import Path
import warnings

# 忽略Warnings
warnings.filterwarnings("ignore")
warnings.simplefilter(action='ignore', category=FutureWarning)

print('''
# 开发测试环境
# Python 3.12.2
# Pandas 2.2.2
# numpy 1.26.4
# matplotlib 3.10.1
# seaborn 0.13.2
# scipy 1.13.0
# statsmodels-0.14.1
\n''')

# 选择数据文件
single_xlsx_fileNames = filedialog.askopenfilename(initialdir="D:/", filetypes=[("xlsx 文件", "*.xlsx")])

if not single_xlsx_fileNames:
    sys.exit("没有可用数据文件，程序终止\n")
    
file_path = Path(single_xlsx_fileNames)
outputFile = file_path.with_name(file_path.stem + '_Output' + file_path.suffix)

# 读取Excel数据
dataXlsFile = single_xlsx_fileNames
_full_dataFrm = pd.read_excel(dataXlsFile, engine='openpyxl')

# 交互方式选择变量
model_var_Obs = []
model_var_Y = []
model_var_X = []

#获取全部变量（列名）
def inputXlsAllVars(_xls_dataFrm):
    print( "\n-->> 可选变量列表：\n" )
    _var_DataFrm = _xls_dataFrm
    _list_dataFm = pd.DataFrame()    
    _list_dataFm["变量"] = _var_DataFrm.columns
    _list_dataFm['变量代码'] = _list_dataFm.index.to_series().apply(lambda x: chr(ord('a') + x)).str.upper()
    _list_dataFm = _list_dataFm.reindex(columns=['变量代码','变量'])
    print(_list_dataFm)
    return _list_dataFm

print("\n-->> 输入对应的变量代码，选择相应变量：\n")

# 选择 样本变量
_for_obs_full_dataFrm = _full_dataFrm
_all_input_vars_DF = inputXlsAllVars(_for_obs_full_dataFrm)

while True:
    _breakExit = False
    my_choice_Obs = input('\n-->> 选择 样本变量：').casefold()
    input_alphabet = re.sub(r'[^a-zA-Z]','', my_choice_Obs)             
    if len(input_alphabet) == 0:
        print("\n-->> 错误！只能选择 字母(a-z)！")
    else:
        input_alphabet = input_alphabet[0]
        input_alphabet_list = []
        input_alphabet_list[:0] = input_alphabet
        input_alphabet_list = list(set(input_alphabet.upper()))
        _testing_obs_DF = _all_input_vars_DF[_all_input_vars_DF['变量代码'].isin(input_alphabet_list)]
        model_var_Obs = _testing_obs_DF['变量'].values.tolist()
        print("\n-->> 样本变量：" + str(model_var_Obs))

        if len(model_var_Obs) > 0:
            obs_choice_GoNext = input('\n-->> 确定（Y/N）:  ').upper()
            if obs_choice_GoNext == 'Y':
                _breakExit = True
    if _breakExit == True:
        break

print('\n-->> 样本变量：' + str(model_var_Obs) + '\n')    

# 选择 被解释变量 Y
_for_yyy_full_dataFrm = _full_dataFrm.drop(model_var_Obs, axis=1)
_all_input_vars_DF = inputXlsAllVars(_for_yyy_full_dataFrm)

while True:
    _breakExit = False
    my_choice_Y = input('\n-->> 选择 被解释变量：').casefold()
    input_alphabet_YY = re.sub(r'[^a-zA-Z]','', my_choice_Y)
    if len(input_alphabet_YY) == 0:
        print("\n-->> 错误！只能选择 字母(a-z)！")
    else:
        input_alphabet_YY = input_alphabet_YY[0]
        input_alphabet_YY_list = []
        input_alphabet_YY_list[:0] = input_alphabet_YY
        input_alphabet_YY_list = list(set(input_alphabet_YY.upper()))
        _testing_YY_DF = _all_input_vars_DF[_all_input_vars_DF['变量代码'].isin(input_alphabet_YY_list)]
        model_var_Y = _testing_YY_DF['变量'].values.tolist()
        print('\n-->> 被解释变量：' + str(model_var_Y)) 
        if len(model_var_Y) > 0:
            yy_choice_GoNext = input('\n-->> 确定（Y/N）:  ').upper()
            if yy_choice_GoNext == 'Y':
                _breakExit = True
    if _breakExit == True:
        break

print('\n-->> 被解释变量：' + str(model_var_Y) + '\n')   
            
# 选择 解释变量 X
_for_xxx_full_dataFrm = _full_dataFrm.drop(model_var_Obs+model_var_Y, axis=1)
_all_input_vars_DF = inputXlsAllVars(_for_xxx_full_dataFrm)
while True:
    _breakExit = False
    my_choice_X = input('\n-->> 选择 解释变量：').casefold()
    input_alphabet_XX = re.sub(r'[^a-zA-Z]','', my_choice_X)
    input_alphabet_XX_list = []
    input_alphabet_XX_list[:0] = input_alphabet_XX
    input_alphabet_XX_list = list(set(input_alphabet_XX.upper()))
    if len(input_alphabet_XX_list) == 0:
        print("\n-->> 错误！只能选择 字母(a-z)！")
    else:
        _testing_XX_DF = _all_input_vars_DF[_all_input_vars_DF['变量代码'].isin(input_alphabet_XX_list)]
        model_var_X = _testing_XX_DF['变量'].values.tolist()
        print('\n-->> 解释变量：' + str(model_var_X)) 
        if len(model_var_X) > 0:
            xx_choice_GoNext = input('\n-->> 确定（Y/N）:  ').upper()
            if xx_choice_GoNext == 'Y':
                _breakExit = True
    if _breakExit == True:
        break

print('\n-->> 解释变量：' + str(model_var_X) + '\n')   

# 0值插补函数
def zero_imputation(data, columns):
    data_copy = data.copy()
    for col in columns:
        if data_copy[col].dtype in ['float64', 'int64']:  # 仅对数值列填充0
            data_copy[col].fillna(0, inplace=True)
        else:  # 非数值列填充"Missing"
            data_copy[col].fillna("Missing", inplace=True)
    return data_copy

# 分析缺失值
print('\n' + '='*20 + ' 缺失值分析 ' + '='*20 + '\n')
missing_summary = _full_dataFrm.isnull().sum()
missing_percentage = (_full_dataFrm.isnull().sum() / len(_full_dataFrm)) * 100
missing_df = pd.DataFrame({'缺失数量': missing_summary, '百分比': missing_percentage})
print(missing_df)
print('\n含有缺失值的行数：', _full_dataFrm.isnull().any(axis=1).sum())
print('总缺失值数量：', _full_dataFrm.isnull().sum().sum())
print('\n' + '='*50 + '\n')

# 选择变量
dataFrm_vars = _full_dataFrm[model_var_Obs + model_var_Y + model_var_X]

# 处理缺失值
output_DF_XLS_List = []
print('\n-->> 处理缺失值\n')
if missing_df['缺失数量'].sum() == 0:
    print('未发现缺失值。')
    dataFrm = dataFrm_vars
else:
    print('检测到缺失值，请选择处理方法：')
    print('1. 删除缺失值')
    print('2. 0值插补')
    print('3. 前向填充 + 后向填充')
    while True:
        choice = input('\n-->> 输入选择 (1/2/3)： ').strip()
        if choice in ['1', '2', '3']:
            break
        print('无效选择，请输入1、2或3。')
    
    # 仅对被解释变量和解释变量进行插补
    imputation_columns = [col for col in model_var_Y + model_var_X if col not in model_var_Obs]
    
    if choice == '1':
        dataFrm = dataFrm_vars.dropna()
        print(f'删除了 {len(dataFrm_vars) - len(dataFrm)} 个缺失值数据\n')
    elif choice == '2':
        dataFrm = zero_imputation(dataFrm_vars, imputation_columns)
        print('已应用0值插补（数值列填充0，非数值列填充"Missing"）\n')
    elif choice == '3':
        dataFrm = dataFrm_vars.ffill().bfill()
        print('已应用 前向填充+后向填充 处理缺失值\n')
    
dataFrm.reset_index(drop=True, inplace=True)
output_DF_XLS_List.append([dataFrm, '数据'])

# 变量关系对图
sns.pairplot(dataFrm, y_vars=model_var_Y, x_vars=[col for col in model_var_X if col in dataFrm.columns], aspect=1)
plt.subplots_adjust(hspace=0.2, wspace=0.2)
plt.title('变量关系图')
plt.show()
plt.close()

# 异常值和影响点分析
print('\n' + '='*20 + ' 异常值和影响点分析 ' + '='*20 + '\n')
print('异常值和影响点的判断标准：\n')
print('1. |标准化残差| > 2（潜在异常值）')
print('2. 杠杆值 > (2k+2)/n（高杠杆点，k=解释变量数，n=样本量）')
print('3. Cook距离 > 4/n（影响点）')
print('4. |DFFITS| > 2*sqrt(k/n)（对拟合值的影响）')
print('\n' + '='*60 + '\n')

# 拟合初始OLS模型以检测异常值
_test_X = dataFrm[model_var_X]
_test_X = sm.add_constant(_test_X)
_test_Y = dataFrm[model_var_Y[0]]
_test_Model = sm.OLS(_test_Y, _test_X, missing='drop').fit()

# 获取影响指标
_outliers = _test_Model.get_influence()
_test_leverage = _outliers.hat_matrix_diag
_test_dffits = _outliers.dffits[0]
_test_resid_stu = _outliers.resid_studentized_external
_test_cook = _outliers.cooks_distance[0]
_test_covratio = _outliers.cov_ratio

# 合并影响指标与样本标识
_outliers_tests = pd.DataFrame({
    '杠杆值': _test_leverage,
    'DFFITS': _test_dffits,
    '标准化残差': _test_resid_stu,
    'Cook距离': _test_cook,
    '协方差比率': _test_covratio
})
_full_outliers = pd.concat([_outliers_tests, dataFrm[model_var_Obs + model_var_Y + model_var_X]], axis=1)

# 定义阈值
_numVar = len(model_var_X)
_numObs = int(_test_Model.nobs)
_cutoff_leverage = (2 * _numVar + 2) / _numObs
_cutoff_cooks = 4 / _numObs
_cutoff_dffits = 2 * math.sqrt(_numVar / _numObs)

# 识别异常值
outliers_resid = _full_outliers[abs(_full_outliers['标准化残差']) > 2]
outliers_leverage = _full_outliers[_full_outliers['杠杆值'] > _cutoff_leverage]
outliers_cooks = _full_outliers[_full_outliers['Cook距离'] > _cutoff_cooks]
outliers_dffits = _full_outliers[abs(_full_outliers['DFFITS']) > _cutoff_dffits]

# 异常值总结
print('-->> 标记异常值总结：\n')
print(f'标准化残差 (|标准化残差| > 2)：{len(outliers_resid)} 个样本\n')
print(outliers_resid[model_var_Obs + model_var_Y + model_var_X])
print(f'\n杠杆值 (杠杆值 > (2k+2)/n = {_cutoff_leverage:.4f})：{len(outliers_leverage)} 个样本\n')
print(outliers_leverage[model_var_Obs + model_var_Y + model_var_X])
print(f'\nCook距离 (Cook距离 > 4/n = {_cutoff_cooks:.4f})：{len(outliers_cooks)} 个样本\n')
print(outliers_cooks[model_var_Obs + model_var_Y + model_var_X])
print(f'\nDFFITS (|DFFITS| > 2*sqrt(k/n) = {_cutoff_dffits:.4f})：{len(outliers_dffits)} 个样本\n')
print(outliers_dffits[model_var_Obs + model_var_Y + model_var_X])
print('\n')

# 识别影响点（Cook距离或DFFITS）
influential_cond = (_full_outliers['Cook距离'] > _cutoff_cooks) | (abs(_full_outliers['DFFITS']) > _cutoff_dffits)
influential_points = _full_outliers[influential_cond][model_var_Obs[0]]

# 识别所有标记异常值
all_outliers_cond = influential_cond | (abs(_full_outliers['标准化残差']) > 2)
all_outliers = _full_outliers[all_outliers_cond][model_var_Obs[0]]

print(f'-->> 总影响点（Cook距离或DFFITS）：{len(influential_points)}')
print(f'-->> 总标记异常值（Cook距离、DFFITS或标准化残差）：{len(all_outliers)}')
print(f'-->> 标记为异常值的数据百分比：{(len(all_outliers) / _numObs) * 100:.2f}%')
print('\n')

# 选择异常值处理方式
print('选择异常值处理方式：')
print('1. 删除影响点（Cook距离或DFFITS）')
print('2. 删除标记异常值（Cook距离、DFFITS或标准化残差）')
print('3. 保留所有数据')
while True:    
    choice = input('\n-->> 输入选择 (1/2/3)： ').strip()
    if choice in ['1', '2', '3']:
        break
    print('无效选择，请输入1、2或3。')

if choice == '1':
    _toDel_cond = dataFrm[model_var_Obs[0]].isin(influential_points)
    _toDel_index = _toDel_cond[_toDel_cond].index
    cleaned_dataFrm = dataFrm.drop(_toDel_index)
    print(f'删除了 {len(_toDel_index)} 个影响点（Cook距离或DFFITS）。')
elif choice == '2':
    _toDel_cond = dataFrm[model_var_Obs[0]].isin(all_outliers)
    _toDel_index = _toDel_cond[_toDel_cond].index
    cleaned_dataFrm = dataFrm.drop(_toDel_index)
    print(f'删除了 {len(_toDel_index)} 个标记异常值（Cook距离、DFFITS或标准化残差）。')
else:
    cleaned_dataFrm = dataFrm
    print('保留所有数据。')

output_DF_XLS_List.append([cleaned_dataFrm, '数据'])

# 清理后数据对图
sns.pairplot(cleaned_dataFrm, y_vars=model_var_Y, x_vars=[col for col in model_var_X if col in cleaned_dataFrm.columns], aspect=1, kind="reg")
plt.subplots_adjust(hspace=0.2, wspace=0.2)
plt.show()
plt.close()

# 为诊断拟合OLS模型
_work_X = cleaned_dataFrm[model_var_X]
_work_X = sm.add_constant(_work_X)
_work_Y = cleaned_dataFrm[model_var_Y[0]]
_work_Model = sm.OLS(_work_Y, _work_X, missing='drop').fit()
_work_model_fitted_Y = _work_Model.fittedvalues
_work_model_residuals = _work_Model.resid

# 描述统计
print('\n' + '='*20 + ' 描述统计 ' + '='*20 + '\n')
descripeStasDF = cleaned_dataFrm.describe(include='all')
descripeStasDF.drop(inplace=True, columns=model_var_Obs, errors='ignore')
descripeStasDF = descripeStasDF.transpose()
descripeStasDF.drop(inplace=True, columns=['25%','50%','75%'], errors='ignore')
descripeStasDF.rename(columns={'count': '样本数', 'freq':'频数', 'mean': '均值', 'std': '标准差', 'min': '最小值', 'max': '最大值'}, inplace=True)
print(descripeStasDF)
output_DF_XLS_List.append([descripeStasDF, '描述统计'])

# 初始化诊断结果列表
diagnostic_results = []

# 线性检验
print('\n' + '='*20 + ' 线性检验 ' + '='*20 + '\n')

# 残差与拟合值图
sns.residplot(x=_work_model_fitted_Y, y=_work_model_residuals, lowess=True, scatter_kws={'alpha': 0.5}, line_kws={'color': 'red', 'lw': 1, 'alpha': 0.8})
plt.title('残差与拟合值图')
plt.xlabel('拟合值')
plt.ylabel('残差')
plt.show()
plt.close()

# Harvey-Collier检验
try:
    __skip = len(_work_Model.params)
    _l_test_Harvey_Collier = sms.recursive_olsresiduals(_work_Model, skip=__skip, alpha=0.95, order_by=None)
    hc_pvalue = stats.ttest_1samp(_l_test_Harvey_Collier[3][__skip:], 0)[1]
    diagnostic_results.append(['Harvey-Collier', hc_pvalue, '线性', hc_pvalue > 0.05])
    print(f'Harvey-Collier p值：{hc_pvalue:.4f}')
except Exception as e:
    print(f'Harvey-Collier检验失败：{e}')
    diagnostic_results.append(['Harvey-Collier', None, '线性', None])

# Rainbow检验
try:
    rainbow_pvalue = linear_rainbow(_work_Model)[1]
    diagnostic_results.append(['Rainbow', rainbow_pvalue, '线性', rainbow_pvalue > 0.05])
    print(f'Rainbow p值：{rainbow_pvalue:.4f}')
except Exception as e:
    print(f'Rainbow检验失败：{e}')
    diagnostic_results.append(['Rainbow', None, '线性', None])

# Ramsey RESET检验
try:
    reset_test = sms.linear_reset(_work_Model, power=2, test_type='fitted')
    reset_pvalue = reset_test.pvalue
    diagnostic_results.append(['Ramsey RESET', reset_pvalue, '线性', reset_pvalue > 0.05])
    print(f'Ramsey RESET p值：{reset_pvalue:.4f}')
except Exception as e:
    print(f'Ramsey RESET检验失败：{e}')
    diagnostic_results.append(['Ramsey RESET', None, '线性', None])

print('结论：如果 p > 0.05 且残差图无明显模式，则支持线性假设。\n')

# 正态性检验
print('\n' + '='*20 + ' 正态性检验 ' + '='*20 + '\n')

# 残差直方图和Q-Q图
sns.histplot(_work_model_residuals, kde=True, stat="density")
plt.xlabel('残差')
plt.title('残差直方图')
plt.show()
plt.close()

try:
    fig = sm.qqplot(_work_model_residuals, dist=stats.t, fit=True, line='45')
    plt.title('Q-Q图')
    plt.show()
    plt.close()
except Exception as e:
    print(f'Q-Q图生成失败：{e}')

# 偏度和峰度
resid_skew = skew(_work_model_residuals)
resid_kurt = kurtosis(_work_model_residuals, fisher=True)
print(f'偏度：{resid_skew:.4f}（接近0表示对称）')
print(f'峰度：{resid_kurt:.4f}（接近0表示正态尾部）')

# Shapiro-Wilk检验
try:
    shapiro_pvalue = shapiro(_work_model_residuals)[1]
    diagnostic_results.append(['Shapiro-Wilk', shapiro_pvalue, '正态性', shapiro_pvalue > 0.05])
    print(f'Shapiro-Wilk p值：{shapiro_pvalue:.4f}')
except Exception as e:
    print(f'Shapiro-Wilk检验失败：{e}')
    diagnostic_results.append(['Shapiro-Wilk', None, '正态性', None])

# Jarque-Bera检验
try:
    jb_pvalue = jarque_bera(_work_model_residuals)[1]
    diagnostic_results.append(['Jarque-Bera', jb_pvalue, '正态性', jb_pvalue > 0.05])
    print(f'Jarque-Bera p值：{jb_pvalue:.4f}')
except Exception as e:
    print(f'Jarque-Bera检验失败：{e}')
    diagnostic_results.append(['Jarque-Bera', None, '正态性', None])

print('结论：如果 p > 0.05，偏度~0，峰度~0，则支持正态性假设。\n')

# 自相关检验
print('\n' + '='*20 + ' 自相关检验 ' + '='*20 + '\n')

# Durbin-Watson检验
dw_stat = durbin_watson(_work_model_residuals)
diagnostic_results.append(['Durbin-Watson', dw_stat, '自相关', 1.5 < dw_stat < 2.5])
print(f'Durbin-Watson：{dw_stat:.4f}（接近2表示无自相关）')

# 自相关函数图
fig = sm.graphics.tsa.plot_acf(_work_model_residuals, lags=min(10, len(_work_model_residuals)//5))
plt.title('残差自相关函数（ACF）图')
plt.show()
plt.close()

# Breusch-Godfrey检验
try:
    bg_lags = min(10, len(_work_model_residuals)//5)
    bg_pvalue = sm.stats.acorr_breusch_godfrey(_work_Model, nlags=bg_lags)[1]
    diagnostic_results.append(['Breusch-Godfrey', bg_pvalue, '自相关', bg_pvalue > 0.05])
    print(f'Breusch-Godfrey p值（滞后={bg_lags}）：{bg_pvalue:.4f}')
except Exception as e:
    print(f'Breusch-Godfrey检验失败：{e}')
    diagnostic_results.append(['Breusch-Godfrey', None, '自相关', None])

print('结论：如果 p > 0.05，Durbin-Watson ~ 2，且ACF无显著滞后，则支持无自相关假设。\n')

# 多重共线性检验
print('\n' + '='*20 + ' 多重共线性检验 ' + '='*20 + '\n')

# VIF计算
_vif_X = cleaned_dataFrm[model_var_X]
_vif_dataFm = pd.DataFrame()
_vif_dataFm["变量"] = _vif_X.columns
_vif_dataFm["VIF"] = [variance_inflation_factor(_vif_X.values, i) for i in range(len(_vif_X.columns))]
diagnostic_results.append(['最大VIF', _vif_dataFm["VIF"].max(), '多重共线性', _vif_dataFm["VIF"].max() < 10])
print(_vif_dataFm)
output_DF_XLS_List.append([_vif_dataFm, 'VIF'])

# 解释变量相关矩阵图
predictor_corr = _vif_X.corr(method='pearson')
sns.heatmap(predictor_corr, cmap='RdBu_r', annot=True, vmin=-1, vmax=1)
plt.title('解释变量相关矩阵图')
plt.show()
plt.close()

# 相关矩阵
correlationsDF = cleaned_dataFrm[model_var_Y + model_var_X].corr(method='pearson')
print('\n相关矩阵：\n')
print(correlationsDF)
output_DF_XLS_List.append([correlationsDF, '相关矩阵'])

print('\n结论：如果VIF < 10 且相关性适中，则多重共线性不显著。\n')

# 异方差检验
print('\n' + '='*20 + ' 异方差检验 ' + '='*20 + '\n')

# Breusch-Pagan检验
try:
    bp_pvalue = het_breuschpagan(_work_model_residuals, _work_X)[1]
    diagnostic_results.append(['Breusch-Pagan', bp_pvalue, '同方差', bp_pvalue > 0.05])
    print(f'Breusch-Pagan p值：{bp_pvalue:.4f}')
except Exception as e:
    print(f'Breusch-Pagan检验失败：{e}')
    diagnostic_results.append(['Breusch-Pagan', None, '同方差', None])

# Goldfeld-Quandt检验
try:
    gq_pvalue = het_goldfeldquandt(_work_model_residuals, _work_X)[1]
    diagnostic_results.append(['Goldfeld-Quandt', gq_pvalue, '同方差', gq_pvalue > 0.05])
    print(f'Goldfeld-Quandt p值：{gq_pvalue:.4f}')
except Exception as e:
    print(f'Goldfeld-Quandt检验失败：{e}')
    diagnostic_results.append(['Goldfeld-Quandt', None, '同方差', None])

print('结论：如果 p > 0.05 且残差图无明显模式，则支持同方差假设。\n')

# 诊断结果总结
print('\n' + '='*20 + ' 诊断结果总结 ' + '='*20 + '\n')
diagnostic_df = pd.DataFrame(diagnostic_results, columns=['检验', '值', '假设', '通过'])
diagnostic_df['结论'] = diagnostic_df.apply(
    lambda row: '通过' if row['通过'] else '未通过' if row['值'] is not None else '未运行', axis=1
)
print(diagnostic_df)
output_DF_XLS_List.append([diagnostic_df, '诊断结果'])

# 回归模型
print('\n' + '='*20 + ' 回归模型 ' + '='*20 + '\n')
ALL_Model_List = []

# OLS模型
ALL_Model_List.append([_work_Model, 'OLS'])
print("=======>>> 基本OLS模型")
print(_work_Model.summary())

# VIF调整后的OLS模型
_over_vif_Vars = _vif_dataFm.loc[_vif_dataFm['VIF'] >= 10]
_keep_vif_Vars = _vif_dataFm.loc[_vif_dataFm['VIF'] < 10]
_over_vif_list = _over_vif_Vars["变量"].values.tolist()
_keep_vif_list = _keep_vif_Vars["变量"].values.tolist()

for over_vif_idx in range(len(_over_vif_list)):
    _vifMdl_name = f'VIF_OLS_{over_vif_idx+1}'
    _vif_X = cleaned_dataFrm[_keep_vif_list + [_over_vif_list[over_vif_idx]]]
    _vif_X = sm.add_constant(_vif_X)
    _vif_Y = cleaned_dataFrm[model_var_Y[0]]
    try:
        _vifMdl = sm.OLS(_vif_Y, _vif_X, missing='drop').fit()
        ALL_Model_List.append([_vifMdl, _vifMdl_name])
    except Exception as e:
        print(f'VIF调整后的OLS模型 ({_vifMdl_name}) 失败：{e}')

# 正则化回归（Ridge、Lasso、Elastic Net）
alphas = 10**np.linspace(10, -2, 100) * 0.01

def min_max_normalize(data):
    return (data - data.min()) / (data.max() - data.min())

_work_X_normalized = _work_X.copy()
for col in model_var_X:
    if col in _work_X_normalized.columns:
        _work_X_normalized[col] = min_max_normalize(_work_X[col])
_work_Y_normalized = min_max_normalize(_work_Y)

def k_fold_cv_regularized(X, y, alphas, L1_wt, k=5):
    n = len(y)
    fold_size = n // k
    mse_scores = []
    for alpha in alphas:
        fold_mse = []
        indices = np.random.permutation(n)
        X_shuffled = X.iloc[indices]
        y_shuffled = y.iloc[indices]
        for i in range(k):
            test_idx = range(i * fold_size, (i + 1) * fold_size)
            train_idx = list(set(range(n)) - set(test_idx))
            X_train = X_shuffled.iloc[train_idx]
            y_train = y_shuffled.iloc[train_idx]
            X_test = X_shuffled.iloc[test_idx]
            y_test = y_shuffled.iloc[test_idx]
            try:
                model = sm.OLS(y_train, X_train).fit_regularized(alpha=alpha, L1_wt=L1_wt)
                y_pred = np.dot(X_test, model.params)
                mse = np.mean((y_test - y_pred) ** 2)
                fold_mse.append(mse)
            except:
                fold_mse.append(np.inf)
        mse_scores.append(np.mean(fold_mse))
    best_alpha_idx = np.argmin(mse_scores)
    return alphas[best_alpha_idx]

try:
    ridge_alpha = k_fold_cv_regularized(_work_X_normalized, _work_Y_normalized, alphas, L1_wt=0)
    lasso_alpha = k_fold_cv_regularized(_work_X_normalized, _work_Y_normalized, alphas, L1_wt=1)
    enet_alpha = k_fold_cv_regularized(_work_X_normalized, _work_Y_normalized, alphas, L1_wt=0.5)

    _sm_pre_mdl = sm.OLS(_work_Y_normalized, _work_X_normalized, missing='drop')
    _sm_pre_mdl_results = _sm_pre_mdl.fit()

    param_names = _work_X_normalized.columns.tolist()

    smfit_ridge = _sm_pre_mdl.fit_regularized(alpha=ridge_alpha, L1_wt=0, start_params=_sm_pre_mdl_results.params)
    ridge_params = pd.Series(smfit_ridge.params, index=param_names)
    ridge_final = sm.regression.linear_model.OLSResults(_sm_pre_mdl, ridge_params, _sm_pre_mdl.normalized_cov_params)

    smfit_lasso = _sm_pre_mdl.fit_regularized(alpha=lasso_alpha, L1_wt=1, start_params=_sm_pre_mdl_results.params)
    lasso_params = pd.Series(smfit_lasso.params, index=param_names)
    lasso_final = sm.regression.linear_model.OLSResults(_sm_pre_mdl, lasso_params, _sm_pre_mdl.normalized_cov_params)

    smfit_elsnet = _sm_pre_mdl.fit_regularized(alpha=enet_alpha, L1_wt=0.5, start_params=_sm_pre_mdl_results.params)
    elsnet_params = pd.Series(smfit_elsnet.params, index=param_names)
    elsnet_final = sm.regression.linear_model.OLSResults(_sm_pre_mdl, elsnet_params, _sm_pre_mdl.normalized_cov_params)

    ALL_Model_List.append([ridge_final, 'Ridge'])
    ALL_Model_List.append([lasso_final, 'Lasso'])
    ALL_Model_List.append([elsnet_final, 'ElasticNet'])
except Exception as e:
    print(f'正则化回归失败：{e}')

# 稳健回归
print('\n' + '='*20 + ' 稳健回归 ' + '='*20 + '\n')
print('详见 https://www.statsmodels.org/stable/rlm_techn1.html\n')

robust_models = [
    ('HuberT', robust_norms.HuberT(), '适度异常值'),
    ('TukeyBiweight', robust_norms.TukeyBiweight(), '严重异常值')
]

for name, norm, description in robust_models:
    try:
        rlm_model = sm.RLM(_work_Y, _work_X, missing='drop', M=norm).fit()
        ALL_Model_List.append([rlm_model, f'稳健{name}'])
        print(f'\n-->> {name} (用途：{description})')
        print(rlm_model.summary())
    except Exception as e:
        print(f'稳健回归 ({name}) 失败：{e}')

# 加权最小二乘（WLS）回归
print('\n' + '='*20 + ' 加权最小二乘（WLS）回归 ' + '='*20 + '\n')

# 权重1：残差方差倒数
resid_var = _work_model_residuals.var()
weights_var = 1 / resid_var
try:
    wls_model_var = sm.WLS(_work_Y, _work_X, weights=weights_var).fit()
    ALL_Model_List.append([wls_model_var, 'WLS（基于方差）'])
    print('\n-->> WLS（基于方差）：残差方差的倒数，稳定样本方差')
    print(wls_model_var.summary())
except Exception as e:
    print(f'WLS（基于方差）失败：{e}')

# 权重2：OLS残差平方的倒数
weights_resid = 1 / (_work_model_residuals ** 2 + 1e-10)
try:
    wls_model_resid = sm.WLS(_work_Y, _work_X, weights=weights_resid).fit()
    ALL_Model_List.append([wls_model_resid, 'WLS（基于残差）'])
    print('\n-->> WLS（基于残差）：OLS残差平方的倒数，针对异方差模式')
    print(wls_model_resid.summary())
except Exception as e:
    print(f'WLS（基于残差）失败：{e}')

# 广义最小二乘（GLS）回归
print('\n' + '='*20 + ' 广义最小二乘（GLS）回归 ' + '='*20 + '\n')

try:
    sigma = np.diag(_work_model_residuals.var() * np.ones(len(_work_Y)))
    gls_model = sm.GLS(_work_Y, _work_X, sigma=sigma).fit()
    ALL_Model_List.append([gls_model, 'GLS'])
    print('\n-->> GLS：估计残差协方差（简单对角矩阵，基于残差方差）处理异方差或自相关')
    print(gls_model.summary())
except Exception as e:
    print(f'GLS回归失败：{e}')

# 分位数回归（Quantile Regression）
print('\n' + '='*20 + ' 分位数回归 ' + '='*20 + '\n')

try:
    quant_model = sm.QuantReg(_work_Y, _work_X).fit(q=0.5)
    ALL_Model_List.append([quant_model, '分位数（中位数）'])
    print('\n-->> 分位数回归（中位数）：建模被解释变量的中位数（0.5或其他分位数）提高对异常值的稳健性')
    print(quant_model.summary())
except Exception as e:
    print(f'分位数回归失败：{e}')

# 诊断检查：WLS和GLS残差的Breusch-Pagan检验
print('\n-->> WLS和GLS模型异方差检查\n')
for model, name in [(wls_model_var, 'WLS基于方差'), (wls_model_resid, 'WLS基于残差'), (gls_model, 'GLS')]:
    try:
        bp_pvalue = sm.stats.diagnostic.het_breuschpagan(model.resid, _work_X)[1]
        print(f'{name} 的Breusch-Pagan p值：{bp_pvalue:.4f}')
        print(f'结论：p > 0.05 表示异方差得到缓解。')
    except Exception as e:
        print(f'{name} 的Breusch-Pagan检验失败：{e}')

# 交叉验证
print('\n' + '='*20 + ' 交叉验证 ' + '='*20 + '\n')
print('执行5折交叉验证以评估模型的样本外均方误差（MSE）。\n')

def k_fold_cv_model(model_obj, X, y, k=5, model_type='OLS', norm=None, **kwargs):
    n = len(y)
    fold_size = n // k
    mse_scores = []
    indices = np.random.permutation(n)
    X_shuffled = X.iloc[indices]
    y_shuffled = y.iloc[indices]
    for i in range(k):
        test_idx = range(i * fold_size, (i + 1) * fold_size)
        train_idx = list(set(range(n)) - set(test_idx))
        X_train = X_shuffled.iloc[train_idx]
        y_train = y_shuffled.iloc[train_idx]
        X_test = X_shuffled.iloc[test_idx]
        y_test = y_shuffled.iloc[test_idx]
        try:
            if model_type == 'OLS':
                model = sm.OLS(y_train, X_train).fit()
            elif model_type == 'RLM':
                model = sm.RLM(y_train, X_train, M=norm).fit()
            elif model_type == 'WLS':
                model = sm.WLS(y_train, X_train, weights=1.0).fit()  # Simplified weights for CV
            elif model_type == 'Regularized':
                model = sm.OLS(y_train, X_train).fit_regularized(alpha=model_obj['alpha'], L1_wt=model_obj['L1_wt'])
            elif model_type == 'GLS':
                sigma = np.diag(np.var(y_train) * np.ones(len(y_train)))
                model = sm.GLS(y_train, X_train, sigma=sigma).fit()
            elif model_type == 'QuantReg':
                model = sm.QuantReg(y_train, X_train).fit(q=0.5)
                
            y_pred = model.predict(X_test)
            mse = np.mean((y_test - y_pred) ** 2)
            mse_scores.append(mse)
        except:
            mse_scores.append(np.inf)
    return np.mean(mse_scores)

cv_results = []
for model, name in ALL_Model_List:
    try:
        if '稳健' in name:
            norm_name = name.split('\n')[1]
            norm = robust_norms.HuberT() if norm_name == 'HuberT' else robust_norms.TukeyBiweight()
            mse = k_fold_cv_model(model, _work_X, _work_Y, model_type='RLM', norm=norm)
        elif 'WLS' in name:
            mse = k_fold_cv_model(model, _work_X, _work_Y, model_type='WLS')
        elif name in ['Ridge', 'Lasso', 'ElasticNet']:
            L1_wt = 0 if name == 'Ridge' else 1 if name == 'Lasso' else 0.5
            alpha = ridge_alpha if name == 'Ridge' else lasso_alpha if name == 'Lasso' else enet_alpha
            mse = k_fold_cv_model({'alpha': alpha, 'L1_wt': L1_wt}, _work_X_normalized, _work_Y_normalized, model_type='Regularized')
        elif name == 'GLS':
            mse = k_fold_cv_model(model, _work_X, _work_Y, model_type='GLS')
        elif '分位数' in name:
            mse = k_fold_cv_model(model, _work_X, _work_Y, model_type='QuantReg')
        else:
            mse = k_fold_cv_model(model, _work_X, _work_Y, model_type='OLS')
        cv_results.append({'模型': name, '交叉验证MSE': mse})
        print(f'{name} 的交叉验证MSE：{mse:.4f}')
    except Exception as e:
        print(f'{name} 的交叉验证失败：{e}')
        cv_results.append({'模型': name, '交叉验证MSE': None})

cv_df = pd.DataFrame(cv_results)
#output_DF_XLS_List.append([cv_df, '交叉验证'])

# 模型比较
print('\n' + '='*20 + ' 模型比较 ' + '='*20 + '\n')
reg_info_dict = {
    '样本数': lambda x: f"{int(x.nobs):d}",
    'AIC': lambda x: f"{x.aic:.4f}" if hasattr(x, 'aic') else 'N/A',
    'BIC': lambda x: f"{x.bic:.4f}" if hasattr(x, 'bic') else 'N/A',
    'F统计量': lambda x: f"{x.fvalue:.4f}" if hasattr(x, 'fvalue') else 'N/A',
    'F检验p值': lambda x: f"{x.f_pvalue:.4f}" if hasattr(x, 'f_pvalue') else 'N/A',
    '伪R²': lambda x: f"{x.prsquared:.4f}" if hasattr(x, 'prsquared') else 'N/A'
}

model_results = []
model_names = []
for model, name in ALL_Model_List:
    model_results.append(model)
    model_names.append(name)

all_regressors = set()
for model, _ in ALL_Model_List:
    all_regressors.update(model.params.index)
all_regressors.discard('const')
regressor_order = [var for var in model_var_X if var in all_regressors] + sorted([var for var in all_regressors if var not in model_var_X])

result_sumTab = summary_col(
    results=model_results,
    float_format='%0.4f',
    stars=True,
    model_names=model_names,
    info_dict=reg_info_dict,
    regressor_order=regressor_order
)
result_sumTab.add_title('回归模型比较')
print(result_sumTab)

# 转换比较表为DataFrame
final_tableDF = result_sumTab.tables[0]
output_DF_XLS_List.append([final_tableDF, '模型比较'])

# 自动模型选择
print('\n' + '='*20 + ' 模型选择 ' + '='*20 + '\n')
print('根据AIC、BIC和交叉验证MSE的综合得分选择最佳模型\n')

model_scores = []
for model, name in ALL_Model_List:
    try:
        aic = model.aic if hasattr(model, 'aic') else np.inf
        bic = model.bic if hasattr(model, 'bic') else np.inf
        mse = next((item['交叉验证MSE'] for item in cv_results if item['模型'] == name), np.inf)
        if np.isinf(aic) or np.isinf(bic) or np.isinf(mse):
            score = np.inf
        else:
            # 标准化指标
            aic_norm = (aic - min([m.aic for m, _ in ALL_Model_List if hasattr(m, 'aic')])) / \
                       (max([m.aic for m, _ in ALL_Model_List if hasattr(m, 'aic')]) - min([m.aic for m, _ in ALL_Model_List if hasattr(m, 'aic')]) + 1e-10)
            bic_norm = (bic - min([m.bic for m, _ in ALL_Model_List if hasattr(m, 'bic')])) / \
                       (max([m.bic for m, _ in ALL_Model_List if hasattr(m, 'bic')]) - min([m.bic for m, _ in ALL_Model_List if hasattr(m, 'bic')]) + 1e-10)
            mse_norm = (mse - min([r['交叉验证MSE'] for r in cv_results if r['交叉验证MSE'] is not None])) / \
                       (max([r['交叉验证MSE'] for r in cv_results if r['交叉验证MSE'] is not None]) - min([r['交叉验证MSE'] for r in cv_results if r['交叉验证MSE'] is not None]) + 1e-10)
            score = (aic_norm + bic_norm + mse_norm) / 3
        model_scores.append({'模型': name, 'AIC': aic, 'BIC': bic, '交叉验证MSE': mse, '综合得分': score})
    except Exception as e:
        print(f'{name} 的模型选择评分失败：{e}')
        model_scores.append({'模型': name, 'AIC': np.inf, 'BIC': np.inf, '交叉验证MSE': np.inf, '综合得分': np.inf})

model_scores_df = pd.DataFrame(model_scores)
model_scores_df = model_scores_df.sort_values('综合得分')
print('\n模型评分（越低越好）：\n')
print(model_scores_df)
best_model = model_scores_df.iloc[0]['模型'] if not np.isinf(model_scores_df.iloc[0]['综合得分']) else '无'
print(f'\n推荐模型：{best_model}')
bestModelRcmdStr = f'\n推荐模型：{best_model}'

output_DF_XLS_List.append([model_scores_df, '模型选择'])

# 写入Excel
print('\n-->> 正在写入Excel...\n')
my_XLS_writer = pd.ExcelWriter(outputFile, engine='openpyxl')

for _df_item in range(len(output_DF_XLS_List)):
    __dfrm = output_DF_XLS_List[_df_item][0]
    __name = output_DF_XLS_List[_df_item][1]
    __dfrm.to_excel(my_XLS_writer, sheet_name = __name, index = True)

# 回归结果添加注释
bottom_row = my_XLS_writer.sheets['模型比较'].max_row
bottom_row2 = my_XLS_writer.sheets['模型选择'].max_row

note_bottom_rowDF_line1 = pd.DataFrame({'注：* p<0.1, ** p<0.05, *** p<0.01, 括号内为标准误。'})
note_bottom_rowDF_line2 = pd.DataFrame({bestModelRcmdStr})

note_bottom_rowDF_line1.to_excel(my_XLS_writer, sheet_name = '模型比较', startrow = bottom_row, index = False, header= False)
note_bottom_rowDF_line2.to_excel(my_XLS_writer, sheet_name = '模型选择', startrow = bottom_row2+1, index = False, header= False)

my_XLS_writer.close()

print('-->> Excel输出完成')

