library(plm)
library(lmtest)
library(texreg)
library(summarytools)
library(openxlsx)

#使用说明 text
infoText <- "　　　　面板数据分析　　　　
1. 仅支持 xlsx 格式数据文件
2. 数据文件格式要求：
　第1列为 样本id
　第2列为 时间time
　第3列为 被解释变量Y
　其他列为 解释变量X
3. 变量名称支持数字、英文字母和汉字"
noticeMsg <- winDialog("ok", infoText)
# 浏览选择 Excel 文件
xlsxFilters <- matrix(c("Excel 文件", "*.xlsx"), 1, 2, byrow = TRUE)
xlsxFilesVec <- choose.files(caption = "数据文件", filters = xlsxFilters, multi = TRUE)
# 数据导入
xlsxFilesNameStr <- xlsxFilesVec[1]
mydata <- read.xlsx(xlsxFilesNameStr)
myColNames <- colnames(mydata)
newColTTTNames <- gsub("\\W", "", myColNames)
newColNames <- gsub(".*^","v_", newColTTTNames)
outPutFileName <- sub("\\.xlsx$", "_Output.xlsx", xlsxFilesNameStr)
colnames(mydata) <- newColNames
myColNames <- colnames(mydata)
index_id <- myColNames[1]
index_time <- myColNames[2]
model_Y <- myColNames[3]
varX_vectors <- myColNames[! myColNames %in% c(index_id,index_time,model_Y)]
# 处理缺失值 替换为 0
mydata[is.na(mydata)] <- 0
# 回归公式	
model_X <- paste(varX_vectors, collapse = "+")
str_reg_formula <- paste(model_Y,model_X, sep=" ~ ", collapse="")
model_formula <- as.formula(str_reg_formula)
# 描述统计
col_to_drops <- c(index_id,index_time)
describe_data <- mydata[ , !(names(mydata) %in% col_to_drops)]
desTable <- descr(describe_data, headings = TRUE, stats = c("mean","med","sd","min","max"), transpose = TRUE)
desTable_re_index <- cbind(rownames(desTable), data.frame(desTable, row.names=NULL))
colnames(desTable_re_index)[1] <- "Variable"
# 模型 result 保存，用于输出文件 
end_list <- list()
end_vect <- c()
# OLS回归
ols_model <- lm(model_formula, data=mydata)
summary(ols_model)
# 混合效应模型
pool_model <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="pooling")
if (exists("pool_model")) {
	end_list[[length(end_list) + 1]] <- pool_model
	end_vect <- append(end_vect, "Pooled")
	summary(pool_model)
} else {
	print("Object does not exist!")
}	
# 固定效应模型
# 固定效应模型 individual
fixed_model_ObsEff <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="within", effect = "individual")
end_list[[length(end_list) + 1]] <- fixed_model_ObsEff
end_vect <- append(end_vect, "Individual Fixed Effects")
summary(fixed_model_ObsEff)
# ★★★ If the p-value is < 0.05 then there is Individual Fixed Effects. 
fixed_obsEff_test <- plmtest(fixed_model_ObsEff, effect="individual")
F_obsEff_pval <- as.numeric(fixed_obsEff_test[["p.value"]][[1]])
# 固定效应模型 time
fixed_model_timeEff <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="within", effect = "time")
end_list[[length(end_list) + 1]] <- fixed_model_timeEff
end_vect <- append(end_vect, "Time Fixed Effects")
summary(fixed_model_timeEff)
# ★★★ If the p-value is < 0.05 then there is Time Fixed Effects. 
fixed_timeEff_test <- plmtest(fixed_model_timeEff, effect="time")
F_timeEff_pval <- as.numeric(fixed_timeEff_test[["p.value"]][[1]])
# 固定效应模型 twoway
fixed_model_twoEff <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="within", effect = "twoways")
end_list[[length(end_list) + 1]] <- fixed_model_twoEff
end_vect <- append(end_vect, "Twoway Fixed Effects")
summary(fixed_model_twoEff)
# 选择 Fixed Effects 模型
if( F_obsEff_pval<0.05 & F_timeEff_pval<0.05 ) {
	fixed_model <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="within", effect = "twoways")
	print("twoway effects")
} else if (F_obsEff_pval<0.05 ) {
	fixed_model <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="within", effect = "individual")
	print("individual effects")
} else if (F_timeEff_pval<0.05 ) {
	fixed_model <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="within", effect = "time")
	print("time effects")
} else {
	fixed_model <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="within")
	print("No Fixed Effect")
}
if (exists("fixed_model")) {
	end_list[[length(end_list) + 1]] <- fixed_model
	end_vect <- append(end_vect, "Fixed Effects")
} 
# 随机效应模型
tryCatch({
	random_model <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="random")
	}, error = function(e) {		
		tryCatch({
			random_model <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="random", random.method = "amemiya")
			}, error = function(e) {
				tryCatch({
					random_model <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="random", random.method = "walhus")
					}, error = function(e) {
					stop("随机效应模型错误！")		
				})		
		})		
})
end_list[[length(end_list) + 1]] <- random_model
end_vect <- append(end_vect, "Random Effects")
summary(random_model)
# 模型选择
# F 检验
pFtest(fixed_model, pool_model) 
model_fixed_pool_test <- pFtest(fixed_model, pool_model)
model_F_test_pval <- as.numeric(model_fixed_pool_test[["p.value"]][[1]])
# Hausman Test
phtest(fixed_model, random_model)
# ★★★ If the p-value is < 0.05 then use fixed effects
model_fixed_random_test <- phtest(fixed_model, random_model)
model_hausman_test_pval <- as.numeric(model_fixed_random_test[["p.value"]][[1]])
# LM 检验
plmtest(pool_model, type=c("bp"))
model_random_pool_test <- plmtest(pool_model, type=c("bp"))
model_LM_test_pval <- as.numeric(model_random_pool_test[["p.value"]][[1]])
# Wooldridge’s first-difference-based test
first_diff_model <- plm(model_formula, data=mydata, index=c(index_id, index_time), model="fd")
if (exists("first_diff_model")) {
	end_list[[length(end_list) + 1]] <- first_diff_model
	end_vect <- append(end_vect, "First-Difference")
	summary(first_diff_model)
} else {
	print("Object does not exist!")
}	
# 模型选择检验结果
model_select_tests <- data.frame(
	检验 = c("F检验","Hausman检验","LM检验"),
	p_value = c(model_F_test_pval,model_hausman_test_pval,model_LM_test_pval), 
	内容 = c("pool vs fixed","fixed vs random","pool vs random"),
	说明 = c("p<0.05选择固定效应","p<0.05选择固定效应","p<0.05选择随机效应"),
	stringsAsFactors = FALSE
)
print(model_select_tests)
#
# 假设检验
#
# Testing for cross-sectional dependence/contemporaneous correlation（截面相关/同期相关）
# Pesaran's CD test 
assump_csd_cd_pval <- NA
tryCatch({ 
	assump_csd_cd_test <- pcdtest(fixed_model_twoEff, test = "cd")
	assump_csd_cd_pval <- as.numeric(assump_csd_cd_test[["p.value"]][[1]])
	}, error = function(e){ 
	assump_csd_cd_pval <- NA
})	
# Scaled Breusch-Pagan LM test 
assump_csd_sclm_pval <- NA
tryCatch({ 
	assump_csd_sclm_test <- pcdtest(fixed_model_twoEff, test = c("sclm"))
	assump_csd_sclm_pval <- as.numeric(assump_csd_sclm_test[["p.value"]][[1]])
	}, error = function(e){ 
	assump_csd_sclm_pval <- NA
})	
# Bias-corrected Scaled Breusch-Pagan LM test 
assump_csd_bcsclm_pval <- NA
tryCatch({ 
	assump_csd_bcsclm_test <- pcdtest(fixed_model_twoEff, test = c("bcsclm"))
	assump_csd_bcsclm_pval <- as.numeric(assump_csd_bcsclm_test[["p.value"]][[1]])
	}, error = function(e){ 
	assump_csd_bcsclm_pval <- NA
})	
	# Testing for serial correlation（序列相关）
assump_sc_fe_pval <- NA
tryCatch({ 
	assump_sc_fe_test <- pbgtest(fixed_model)
	assump_sc_fe_pval <- as.numeric(assump_sc_fe_test[["p.value"]][[1]])
	}, error = function(e){ 
	assump_sc_fe_pval <- NA
})		
assump_sc_pe_pval <- NA
tryCatch({ 
	assump_sc_pe_test <- pbgtest(pool_model)
	assump_sc_pe_pval <- as.numeric(assump_sc_pe_test[["p.value"]][[1]])
	}, error = function(e){ 
	assump_sc_pe_pval <- NA
})		
assump_sc_re_pval <- NA
tryCatch({ 
	assump_sc_re_test <- pbgtest(random_model)
	assump_sc_re_pval <- as.numeric(assump_sc_re_test[["p.value"]][[1]])
	}, error = function(e){ 
	assump_sc_re_pval <- NA
})			
# Testing for unit roots/stationarity（单位根/平稳性）
mydata.set <- pdata.frame(mydata, index = c(index_id, index_time))
txt_adf_test <- paste("mydata.set$",model_Y, sep="", collapse="")
exp_adf_test <- eval(parse(text = txt_adf_test))
assump_adf_pval <- NA
tryCatch({ 
	assump_adf_test <- adf.test(exp_adf_test, k=2)
	assump_adf_pval <- as.numeric(assump_adf_test[["p.value"]][[1]])
	}, error = function(e){ 
	assump_adf_pval <- NA
})	
assump_adf_madwu_pval <- NA
tryCatch({ 
	assump_adf_madwu_test <- purtest(exp_adf_test, pmax = 4, exo = "intercept", test = "madwu")
	assump_adf_madwu_pval <- as.numeric(assump_adf_madwu_test[["statistic"]][["p.value"]][[1]])
	}, error = function(e){ 
	assump_adf_madwu_pval <- NA
})	
assump_adf_Pm_pval <- NA
tryCatch({ 
	assump_adf_Pm_test <- purtest(exp_adf_test, pmax = 4, exo = "intercept", test = "Pm")
	assump_adf_Pm_pval <- as.numeric(assump_adf_Pm_test[["statistic"]][["p.value"]][[1]])
	}, error = function(e){ 
	assump_adf_Pm_pval <- NA
})	
# Cross-sectionally Augmented IPS Test for Unit Roots in Panel Models
assump_adf_CIPS_pval <- NA
tryCatch({ 
	assump_adf_CIPS_test <- cipstest(exp_adf_test, type = "trend")
	assump_adf_CIPS_pval <- as.numeric(assump_adf_CIPS_test[["p.value"]][[1]])
	}, error = function(e){ 
	assump_adf_CIPS_pval <- NA
})
# The Hadri Lagrange multiplier (LM) test 
assump_adf_Hadri_pval <- NA
tryCatch({ 
	assump_adf_Hadri_test <- purtest(exp_adf_test, exo = "intercept", test = "hadri")
	assump_adf_Hadri_pval <- as.numeric(assump_adf_Hadri_test[["statistic"]][["p.value"]][[1]])
	}, error = function(e){ 
	assump_adf_Hadri_pval <- NA
})	
# Im-Pesaran-Shin Unit-Root Test 
assump_IPS_pval <- NA
tryCatch({ 
	assump_IPS_test <- purtest(exp_adf_test, pmax = 4, exo = "intercept", test = "ips")
	assump_IPS_pval <- as.numeric(assump_IPS_test[["statistic"]][["p.value"]][[1]])
	}, error = function(e){ 
	assump_IPS_pval <- NA
})
# Levin-Lin-Chu Unit-Root Test 
assump_Levinlin_pval <- NA
tryCatch({ 
	assump_Levinlin_test <- purtest(exp_adf_test, pmax = 4, exo = "intercept", test = "levinlin")
	assump_Levinlin_pval <- as.numeric(assump_Levinlin_test[["statistic"]][["p.value"]][[1]])
	}, error = function(e){ 
	assump_Levinlin_pval <- NA
})
# Testing for heteroskedasticity（异方差）
str_hete_test <- paste(str_reg_formula,"+factor(",index_id,")", sep="", collapse="")
model_hete_test <- as.formula(str_hete_test)
assump_hete_pval <- NA
tryCatch({ 
	assump_hete_test <- bptest(model_hete_test, data = mydata, studentize=F)
	assump_hete_pval <- as.numeric(assump_hete_test[["p.value"]][[1]])
	}, error = function(e){ 
	assump_hete_pval <- NA
})		
# 假设检验结果
assumptions_tests <- data.frame(
	假设检验 = c("cross-sectional dependence"," "," ","serial correlation"," "," ","unit roots"," "," "," ","heteroskedasticity"),
	方法 = c("Pesaran CD test","Breusch-Pagan LM test","Bias-corrected Scaled Breusch-Pagan LM test","Breusch-Godfrey/Wooldridge test( fixed model )","Breusch-Godfrey/Wooldridge test ( pool model )","Breusch-Godfrey/Wooldridge test ( random model )","Augmented Dickey-Fuller Test","Pesaran's CIPS test","Levin-Lin-Chu Unit-Root Test","Hadri Test","Breusch-Pagan test"), 
	p_value = c(assump_csd_cd_pval,assump_csd_sclm_pval,assump_csd_bcsclm_pval,assump_sc_fe_pval,assump_sc_pe_pval,assump_sc_re_pval,assump_adf_pval,assump_adf_CIPS_pval,assump_Levinlin_pval,assump_adf_Hadri_pval,assump_hete_pval), 
	说明 = c("p<0.05即有","p<0.05即有","p<0.05即有","p<0.05即有","p<0.05即有","p<0.05即有","p<0.05即没有","p<0.05即没有","p<0.05即没有","p<0.05即有","p<0.05即有"),
	stringsAsFactors = FALSE
)
print(assumptions_tests)
# Robust covariance matrix estimation
fixed_model_hac <- coeftest(fixed_model, vcovHC(fixed_model, method = "arellano", type = "sss"))
if (exists("fixed_model_hac")) {
	end_list[[length(end_list) + 1]] <- fixed_model_hac
	end_vect <- append(end_vect, "Fixed Model Robust")
	summary(fixed_model_hac)
} 	
tryCatch({ 
	random_model_hac3 <- coeftest(random_model, vcovHC(random_model))
	}, error = function(e){ 
	random_model_hac3 <- NA
	}, warning = function(w){
	random_model_hac3 <- NA
})	
if (exists("random_model_hac3")) {
	end_list[[length(end_list) + 1]] <- random_model_hac3
	end_vect <- append(end_vect, "Random Model Robust")
	summary(random_model_hac3)
} 	
# Beck and Katz (1995) method or Panel Corrected Standard Errors (PCSE)
fixed_model_pcse <- coeftest(fixed_model, vcovBK(fixed_model, type="HC3", cluster = "group")) 
if (exists("fixed_model_pcse")) {
	end_list[[length(end_list) + 1]] <- fixed_model_pcse
	end_vect <- append(end_vect, "Fixed Model PCSE(robust vs. serial correlation)")
	summary(fixed_model_pcse)
} 	
fixed_model_time_pcse <- coeftest(fixed_model, vcovBK(fixed_model, type="HC3", cluster = "time")) 
if (exists("fixed_model_time_pcse")) {
	end_list[[length(end_list) + 1]] <- fixed_model_time_pcse
	end_vect <- append(end_vect, "Fixed Model PCSE(robust vs. cross-sectional correlation)")
	summary(fixed_model_time_pcse)
}	
# Driscoll and Kraay standard errors /spatial correlation consistent standard errors/SCC
fixed_model_scc <- coeftest(fixed_model, vcovSCC(fixed_model, type="HC3"))
if (exists("fixed_model_scc")) {
	end_list[[length(end_list) + 1]] <- fixed_model_scc
	end_vect <- append(end_vect, "Fixed Model SCC(for panel models with cross-sectional and serial correlation)")
	summary(fixed_model_scc)
}	
# vcovDC is a function for estimating a robust covariance matrix of parameters for a panel model with errors clustering along both dimensions. 
fixed_model_vcovDC <- coeftest(fixed_model, vcovDC(fixed_model, type="HC3"))
if (exists("fixed_model_vcovDC")) {
	end_list[[length(end_list) + 1]] <- fixed_model_vcovDC
	end_vect <- append(end_vect, "Fixed Model Double-Clustering Robust")
}
# 回归结果表格
for (modelName in end_list){
	print(modelName)
}
for (modelNote in end_vect ){
	print(modelNote)
}
screenreg(end_list, 
		  custom.model.names = end_vect,
		  stars = c(0.01, 0.05, 0.1),
		  custom.note = "%stars.  Standard errors in parentheses." 
		  )
modelResult <- matrixreg(end_list, 
		  custom.model.names = end_vect,
		  stars = c(0.01, 0.05, 0.1),
		  trim = FALSE,
		  custom.note = "%stars.  Standard errors in parentheses." 
		  )
# 添加note
modelResultDF <- data.frame(modelResult)
modelResultDF <- rbind(modelResultDF, " ") 
modelResultDF[nrow(modelResultDF),1] <- "Standard errors in parentheses.  *** p < 0.01; ** p < 0.05; * p < 0.1."
# xlsx 输出
xls_output_list <- list('数据'=mydata, '描述统计'=desTable_re_index, '模型选择'=model_select_tests,'假设检验'=assumptions_tests, '模型结果'=modelResultDF)
write.xlsx(xls_output_list, file = outPutFileName, startCol = 1, startRow = 1, rowNames=FALSE, asTable=FALSE, overwrite=TRUE)
print("----------------- 程序结束 ---------------------")
