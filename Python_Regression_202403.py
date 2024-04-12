# -*- coding: utf-8 -*-

import os
import sys
import re
import warnings
warnings.filterwarnings('ignore')
warnings.simplefilter(action='ignore', category=FutureWarning)
from tkinter import filedialog
import pandas as pd
from openpyxl import Workbook
import numpy as np
from scipy import stats
import statsmodels.api as sm
import statsmodels.stats.api as sms
from statsmodels.tsa.stattools import adfuller, kpss
from statsmodels.stats.api import linear_harvey_collier
from statsmodels.stats.diagnostic import linear_rainbow
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.diagnostic import normal_ad
from statsmodels.stats.diagnostic import kstest_normal 
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.diagnostic import acorr_ljungbox, acorr_breusch_godfrey  
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.stats.diagnostic import het_breuschpagan  
from statsmodels.stats.api import het_goldfeldquandt  
from statsmodels.stats.diagnostic import het_white
from statsmodels.iolib.summary2 import summary_col

##################################################
#
# python Batch Regression
#
##################################################

dataMultiFileList = []
output_DF_XLS_List = []

model_var_Obs = []
model_var_Y = []
model_var_X = []

# 继续执行或终止
def userChoiceCtrl():
    while True:
        my_choice = input('\n-->> 是否继续执行，请选择 Y/N ： ').casefold()
        if my_choice == "y":
            break
        elif my_choice == "n":
            sys.exit('用户终止')

print('''\n
##################################################
#                                                #
#           OLS Regression 程序开始              #
#                                                #
##################################################
\n''')

print('''
# 开发测试环境
# Python 3.12.2
# Pandas 2.2.2
# openpyxl-3.1.2
# statsmodels-0.14.1
\n''')
print('''
------ 数据文件格式要求 ------

  1. 仅支持 xlsx 格式
  2. 工作表第一列为样本变量数据
  3. 工作表第二列为被解释变量数据
  4. 工作表其他列为解释变量数据
  5. 缺失数据将自动替换为前值
  
------------------------------
\n''')
userChoiceCtrl()
# 选择数据文件
all_xlsx_fileNames = filedialog.askopenfilenames(initialdir="D:/", filetypes=[("xlsx 文件", "*.xlsx")])
if all_xlsx_fileNames:
    for _filename in all_xlsx_fileNames:
        one_xlsx_filename = _filename.replace("/", "\\")
        dataMultiFileList.append(one_xlsx_filename)
else:
    sys.exit("没有可用数据文件，程序终止\n")

for _xls_name in dataMultiFileList:
    output_DF_XLS_List = []
    filename, file_extension = os.path.splitext(_xls_name)
    outputFile = filename + '_Output' + file_extension
    # 读取 xls 数据
    dataXlsFile = _xls_name
    _full_dataFrm = pd.read_excel(dataXlsFile)    
    _full_dataFrm.columns = ['v_' + re.sub(r"\W", "",str(col)) for col in _full_dataFrm.columns]
    _full_dataFrm_cols_list = _full_dataFrm.columns.tolist()
    if len(_full_dataFrm_cols_list) < 3:
        print("错误：　" + _xls_name + "　变量不足")
        pass
    else:
        # 变量
        model_var_Obs = _full_dataFrm_cols_list[0:1]
        model_var_Y = _full_dataFrm_cols_list[1:2]
        model_var_X = _full_dataFrm_cols_list[2:]    
    # 缺失值处理
    dataFrm_vars = _full_dataFrm.ffill()
    cleaned_dataFrm = dataFrm_vars.bfill()
    output_DF_XLS_List.append([cleaned_dataFrm,'数据'])
    # 基本模型
    _work_X = cleaned_dataFrm[model_var_X]
    _work_Y = cleaned_dataFrm[model_var_Y[0]]
    _work_Model = sm.OLS(_work_Y, sm.add_constant(_work_X), missing='drop').fit()
    _work_model_fitted_Y = _work_Model.fittedvalues
    _work_model_residuals = _work_Model.resid
    # 描述统计
    descripeStasDF = cleaned_dataFrm.describe(include='all')
    descripeStasDF.drop(inplace=True, columns=model_var_Obs)
    descripeStasDF = descripeStasDF.transpose()
    descripeStasDF.drop(inplace=True, columns=['25%','50%','75%'])
    descripeStasDF.rename(columns={'count': '样本数', 'freq':'频数', 'mean': '均值', 'std': '标准差', 'min': '最小值', 'max': '最大值'}, inplace=True)
    output_DF_XLS_List.append([descripeStasDF,'描述统计'])
    # 相关系数
    correlationsDF = cleaned_dataFrm[model_var_Y+model_var_X].corr(method ='pearson')
    #correlationsDF = cleaned_dataFrm[model_var_Y+model_var_X].corr(method ='spearman')
    #correlationsDF = cleaned_dataFrm[model_var_Y+model_var_X].corr(method ='kendall')
    output_DF_XLS_List.append([correlationsDF,'相关系数'])
    # 假设检验
    test_resultDF = pd.DataFrame()
    mean_residuals = np.mean(_work_model_residuals)
    # Harvey-Collier multiplier test
    __skip = len(_work_Model.params)
    _l_test_Harvey_Collier = sms.recursive_olsresiduals(_work_Model, skip=__skip, alpha=0.95, order_by=None)
    stats.ttest_1samp(_l_test_Harvey_Collier[3][__skip:], 0)
    hc_pvalue = stats.ttest_1samp(_l_test_Harvey_Collier[3][__skip:], 0)[1]
    # Rainbow test
    rainbow_pvalue = linear_rainbow(_work_Model)[1]
    test_result_data = {
        "假设检验": ["Linear/线性", "", ""],
        "结果": ["★ 残差均值/Mean of Residuals： {}".format("%f"%mean_residuals), '★ p-value of Harvey-Collier test： {}'.format(hc_pvalue), '★ p-value of Rainbow test： {}'.format(rainbow_pvalue)]
        }
    test_Linear_df = pd.DataFrame(test_result_data)
    test_resultDF = pd.concat([test_resultDF, test_Linear_df], ignore_index=True)
    # Shapiro-Wilk test
    shapiro_test_pvalue = stats.shapiro(_work_model_residuals)[1]
    # D'Agostino's K-squared test
    normal_test_pvalue = stats.normaltest(_work_model_residuals)[1]
    # Kolmogorov-Smirnov test
    kstest_normal_pvalue = kstest_normal(_work_model_residuals)[1]
    # Anderson-Darling test
    normal_ad_pvalue = normal_ad(_work_model_residuals)[1]
    test_result_data = {
        "假设检验": ["Normally distributed/正态性", "", "",""],
        "结果": ['★ p-value of Shapiro-Wilk test： {}'.format("%f"%shapiro_test_pvalue), "★ p-value of D'Agostino's K-squared test： {}".format("%f"%normal_test_pvalue), '★ p-value of Kolmogorov-Smirnov test： {}'.format("%f"%kstest_normal_pvalue),'★ p-value of Anderson-Darling test： {}'.format("%f"%normal_ad_pvalue)]
        }
    test_Normally_df = pd.DataFrame(test_result_data)
    test_resultDF = pd.concat([test_resultDF, test_Normally_df], ignore_index=True)
    # Durbin-Watson
    durbin_Watson = durbin_watson(_work_model_residuals)
    # Breusch-Godfrey Lagrange Multiplier tests
    try:
        _BG_lags = sm.stats.acorr_breusch_godfrey(_work_Model, nlags=40)
    except ValueError:
        pass
    test_result_data = {
        "假设检验": ["Independent/No Autocorrelation/独立性/无自相关性"],
        "结果": ['★ Durbin-Watson： {}'.format("%f"%durbin_Watson)]
        }
    test_Independent_df = pd.DataFrame(test_result_data)
    test_resultDF = pd.concat([test_resultDF, test_Independent_df], ignore_index=True)
    # Breusch-Pagan test
    _bp_test_X = cleaned_dataFrm[model_var_X]
    het_breuschpagan_pvalue = sms.het_breuschpagan(_work_model_residuals, _work_Model.model.exog)
    # White general test
    try:
        _white_test_X = cleaned_dataFrm[model_var_X]
        het_white_pvalue = sms.het_white(_work_model_residuals, _work_Model.model.exog)[1]
    except AssertionError:
        pass
    # Goldfeld-Quandt test
    _gq_test_X = cleaned_dataFrm[model_var_X]
    het_goldfeldquandt_pvalue = sms.het_goldfeldquandt(_work_model_residuals, _work_Model.model.exog)[1]
    test_result_data = {
        "假设检验": ["Equal variances/同方差", ""],
        "结果": ['★ LM statistic p-value of Breusch-Pagan test： {}'.format("%f"%het_breuschpagan_pvalue[1]), '★ p-value of Goldfeld-Quandt test： {}'.format("%f"%het_goldfeldquandt_pvalue)]
        }
    test_Equal_df = pd.DataFrame(test_result_data)
    test_resultDF = pd.concat([test_resultDF, test_Equal_df], ignore_index=True)
    output_DF_XLS_List.append([test_resultDF,'假设检验'])
    # VIF
    _vif_X = cleaned_dataFrm[model_var_X]
    _vif_dataFm = pd.DataFrame()
    _vif_dataFm["解释变量 X"] = _vif_X.columns
    _vif_dataFm["VIF"] = [variance_inflation_factor(_vif_X.values, i) for i in range(len(_vif_X.columns))]
    output_DF_XLS_List.append([_vif_dataFm,'VIF'])
    ALL_Model_List = []
    ALL_Model_List.append([_work_Model,'OLS'])
    # 依次单独加入VIF>10的解释变量X，多个模型
    _over_vif_Vars = _vif_dataFm.loc[_vif_dataFm['VIF'] >= 10]
    _keep_vif_Vars = _vif_dataFm.loc[_vif_dataFm['VIF'] < 10]
    _over_vif_list = _over_vif_Vars["解释变量 X"].values.tolist()
    _keep_vif_list = _keep_vif_Vars["解释变量 X"].values.tolist()
    for over_vif_idx in range(len(_over_vif_list)):
        _vifMdl_name = ('vif_OLS_' + str(over_vif_idx+1))
        _vif_X = cleaned_dataFrm[_keep_vif_list+[_over_vif_list[over_vif_idx]]]
        _vif_X = sm.add_constant(_vif_X)
        _vif_Y = cleaned_dataFrm[model_var_Y[0]]
        _vifMdl = sm.OLS(_vif_Y, _vif_X, missing='drop').fit()
        ALL_Model_List.append([_vifMdl,_vifMdl_name])
    # 稳健回归 Robust Regression
    _work_Model_rlm = sm.RLM(_work_Y, _work_X, missing='drop', M=sm.robust.norms.HuberT()).fit()
    ALL_Model_List.append([_work_Model_rlm,'Robust\nHubers t'])
    _work_Model_rlm2 = sm.RLM(_work_Y, _work_X, missing='drop', M=sm.robust.norms.AndrewWave()).fit()
    ALL_Model_List.append([_work_Model_rlm2,'Robust\nAndrew Wave'])
    _work_Model_rlm3 = sm.RLM(_work_Y, _work_X, missing='drop', M=sm.robust.norms.LeastSquares()).fit()
    ALL_Model_List.append([_work_Model_rlm3,'Robust\nLeast Squares'])
    _work_Model_rlm4 = sm.RLM(_work_Y, _work_X, missing='drop', M=sm.robust.norms.TukeyBiweight()).fit()
    ALL_Model_List.append([_work_Model_rlm4,'Robust\nTukeyBiweight'])
    _work_Model_rlm5 = sm.RLM(_work_Y, _work_X, missing='drop', M=sm.robust.norms.RamsayE()).fit()
    ALL_Model_List.append([_work_Model_rlm5,'Robust\nRamsay Ea'])
    # 加权最小二乘法回归
    weight_1 = 1/np.abs(_work_model_fitted_Y)
    model_wls_1 = sm.WLS(_work_Y, _work_X, weights = weight_1).fit()
    ALL_Model_List.append([model_wls_1,'WLS\n拟合值的倒数'])
    wls_cleaned_dataFrm = cleaned_dataFrm.copy()
    wls_cleaned_dataFrm["tmp_weight"] = np.abs(_work_model_residuals)
    _wls_X = wls_cleaned_dataFrm[model_var_X]
    _wls_X = sm.add_constant(_wls_X)
    _wls_weight_Y = wls_cleaned_dataFrm["tmp_weight"]
    _wls_Y = wls_cleaned_dataFrm[model_var_Y[0]]
    model_wls_tmp = sm.OLS(_wls_weight_Y, _wls_X, missing='drop').fit()
    weight_2 = model_wls_tmp.fittedvalues
    weight_2 = weight_2**-2
    wls_cleaned_dataFrm['weight_2'] = weight_2
    model_wls_2 = sm.WLS(_wls_Y, _wls_X, wls_cleaned_dataFrm['weight_2']).fit()
    ALL_Model_List.append([model_wls_2,'WLS\n残差平方的倒数'])
    wls_cleaned_dataFrm['logerr2'] = np.log(_work_model_residuals**2)
    _wls_weight_YY = wls_cleaned_dataFrm["logerr2"]
    model_wls_tmpWW = sm.OLS(_wls_weight_YY, _wls_X, missing='drop').fit()
    wls_cleaned_dataFrm['weight_3'] = 1/(np.exp(model_wls_tmpWW.predict()))
    model_wls_3 = sm.WLS(_work_Y, _work_X, weights = wls_cleaned_dataFrm['weight_3']).fit()
    ALL_Model_List.append([model_wls_3,'WLS\n用残差平方的对数重新拟合，拟合值指数的倒数'])
    # 合并输出回归结果
    reg_info_dict={'Observations' : lambda x: f"{int(x.nobs):d}",
                   'AIC' : lambda x: f"{x.aic:.4f}",
                   'BIC' : lambda x: f"{x.bic:.4f}",
                   'F-statistic' : lambda x: f"{x.fvalue:.4f}",
                   'Prob > F' : lambda x: f"{x.f_pvalue:.4f}"
                   }
                   
    model_results = []
    model___names = []
    for _model_idx in range(len(ALL_Model_List)):
        model_results.append(ALL_Model_List[_model_idx][0])
        model___names.append(ALL_Model_List[_model_idx][1])
        
    result_sumTab = summary_col(results=model_results,
                                float_format='%0.4f',
                                stars = True,
                                model_names=model___names,
                                info_dict=reg_info_dict,
                                regressor_order=model_var_X)
    result_sumTab.add_title('回归模型')
    
    final_tableDF = result_sumTab.tables[0]
    output_DF_XLS_List.append([final_tableDF,'回归结果'])
    print('\n正在写入数据文件...　' + outputFile)
    my_XLS_writer = pd.ExcelWriter(outputFile, engine='openpyxl')
    for _df_item in range(len(output_DF_XLS_List)):
        __dfrm = output_DF_XLS_List[_df_item][0]
        __name = output_DF_XLS_List[_df_item][1]
        __dfrm.to_excel(my_XLS_writer, sheet_name = __name, index = True)
    # 添加注释
    bottom_row = my_XLS_writer.sheets['回归结果'].max_row
    note_bottom_rowDF_line1 = pd.DataFrame({'注：Standard errors in parentheses.'})
    note_bottom_rowDF_line2 = pd.DataFrame({'    * p<0.1, ** p<0.05, ***p<0.01'})
    note_bottom_rowDF_line1.to_excel(my_XLS_writer, sheet_name = '回归结果', startrow = bottom_row, index = False, header= False)
    note_bottom_rowDF_line2.to_excel(my_XLS_writer, sheet_name = '回归结果', startrow = bottom_row+1, index = False, header= False)
    my_XLS_writer.close()

print('''
##################################################
#                                                #
#                   程序结束                     #
#                                                #
##################################################
''')



