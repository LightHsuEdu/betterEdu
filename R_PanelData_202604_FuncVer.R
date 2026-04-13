library(plm)
library(lmtest)
library(texreg)
library(summarytools)
library(openxlsx)
library(fixest)
library(tseries)
library(sandwich)

myPanelDataAnalysis_Function <- function(input_xlsx_path) {
  # 辅助函数
  fmt4 <- function(x) ifelse(is.na(x), "", format(round(x, 4), nsmall = 4, trim = TRUE))
  fmt2 <- function(x) ifelse(is.na(x), "", format(round(x, 2), nsmall = 2, trim = TRUE))
  fmt0 <- function(x) ifelse(is.na(x), "", as.character(x))

  make_sig_star <- function(p) {
    if (is.na(p)) return("")
    if (p < 0.01) return("***")
    if (p < 0.05) return("**")
    if (p < 0.1)  return("*")
    return("")
  }

  make_paperStyle_row <- function(values, col_names) {
    if (length(values) < length(col_names)) values <- c(values, rep("", length(col_names) - length(values)))
    if (length(values) > length(col_names)) values <- values[1:length(col_names)]
    row_df <- as.data.frame(matrix(values, nrow = 1), stringsAsFactors = FALSE)
    names(row_df) <- col_names
    row_df
  }

  safe_fit <- function(expr, envir = parent.frame()) {
    tryCatch(
      suppressWarnings(eval(expr, envir = envir)),
	  error = function(e) {
	    message("模型未成功: ", conditionMessage(e))
	    NULL
	  }
	)
  }

  esc_html <- function(x) {
    x <- as.character(x)
    x <- gsub("&", "&amp;", x, fixed = TRUE)
    x <- gsub("<", "&lt;", x, fixed = TRUE)
    x <- gsub(">", "&gt;", x, fixed = TRUE)
    x
  }

  text_log_block <- function(title, desc = NULL, obj = NULL) {
    lines <- if (is.null(obj)) {
      "无输出内容"
    } else {
      capture.output(obj)
    }
    i_call <- which(grepl("^Call:", lines))
    if (length(i_call) == 1L) {
      j <- i_call + 1L
      while (j <= length(lines) &&
        (j == i_call + 1L || grepl("^\\s", lines[j])) &&
        nzchar(lines[j])) {
        j <- j + 1L
      }
      if (j > i_call + 1L) {
        call_lines <- lines[(i_call + 1L):(j - 1L)]
        call_one <- paste(trimws(call_lines), collapse = " ")
        lines <- c(
          lines[1:i_call],
          paste0("  ", call_one),
          if (j <= length(lines)) lines[j:length(lines)] else character(0)
        )
      }
    }
    i_data <- which(grepl("^data:", lines))
    if (length(i_data) >= 1L) {
      for (idx in rev(i_data)) {
        j <- idx + 1L
        while (j <= length(lines) &&
          (grepl("^\\s", lines[j]) || grepl("^structure\\(", lines[j])) &&
          nzchar(lines[j])) {
          j <- j + 1L
        }
        if (j > idx + 1L) {
          lines <- c(
            lines[1:idx],
            "  [data omitted]",
            if (j <= length(lines)) lines[j:length(lines)] else character(0)
          )
        }
      }
    }
    i_formula <- which(grepl("^formula:", lines))
    if (length(i_formula) == 1L) {
      j <- i_formula + 1L
      while (j <= length(lines) &&
        grepl("^\\s", lines[j]) &&
        nzchar(lines[j]) &&
        !grepl("^\\s*(data|link|threshold|nobs)", lines[j])) {
        j <- j + 1L
      }
      if (j > i_formula + 1L) {
        formula_lines <- lines[(i_formula + 1L):(j - 1L)]
        formula_one <- paste(trimws(formula_lines), collapse = " ")
        lines <- c(
          lines[1:i_formula],
          paste0("  ", formula_one),
          if (j <= length(lines)) lines[j:length(lines)] else character(0)
        )
      }
    }
    txt <- esc_html(paste(lines, collapse = "\n"))
    paste0(
      "<div class='log-block'>",
      "<div>========================================================================================================================</div>",
      "<h4 class='log-title'>", esc_html(title), "</h4>",
      if (!is.null(desc)) sprintf("<p class='log-desc'>说明：%s</p>", esc_html(desc)) else "",
      "<div>------------------------------------------------------------</div>",
      "<pre class='log-text'>", txt, "</pre>",
      "</br>",
      "</div>"
    )
  }

  df_to_html_block <- function(df, title = NULL, desc = NULL) {
    if (is.null(df)) df <- data.frame(说明 = "无输出内容", stringsAsFactors = FALSE)
    if (!is.data.frame(df)) {
      df <- tryCatch(as.data.frame(df), error = function(e) {
        data.frame(说明 = "对象无法转换为数据框", stringsAsFactors = FALSE)
      })
    }
    if (nrow(df) == 0) df <- data.frame(说明 = "无数据", stringsAsFactors = FALSE)
    header <- paste0(
      "<tr>",
      paste(sprintf("<th>%s</th>", esc_html(colnames(df))), collapse = ""),
      "</tr>"
    )
    rows <- apply(df, 1, function(r) {
      paste0(
        "<tr>",
        paste(sprintf("<td>%s</td>", esc_html(r)), collapse = ""),
        "</tr>"
      )
    })
    paste0(
      "<div class='log-table-block'>",
      "<div>========================================================================================================================</div>",
      if (!is.null(title)) sprintf("<h4 class='log-title'>%s</h4>", esc_html(title)) else "",
      if (!is.null(desc)) sprintf("<p class='log-desc'>说明：%s</p>", esc_html(desc)) else "",
      "<div>------------------------------------------------------------</div>",
      "<div class='log-table-wrapper'>",
      "<table class='log-table'>",
      header,
      paste(rows, collapse = ""),
      "</table>",
      "</div>",
      "</br>",
      "</div>"
    )
  }

  info_table_to_html_block <- function(info_list, title = "分析说明", desc = NULL) {
    if (is.null(info_list) || length(info_list) == 0) {
      df <- data.frame(项目 = "说明", 内容 = "无可输出内容", stringsAsFactors = FALSE)
    } else {
      df <- data.frame(项目 = names(info_list), 内容 = unname(as.character(info_list)), stringsAsFactors = FALSE)
    }
    df_to_html_block(df, title = title, desc = desc)
  }

  judge_text <- function(x) ifelse(is.na(x), "无法判断", ifelse(x, "通过", "未通过"))
  pass_if_p_greater_005 <- function(p) ifelse(is.na(p), NA, p >= 0.05)
  pass_if_p_less_005 <- function(p) ifelse(is.na(p), NA, p < 0.05)

  combine_log_blocks <- function(...) {
    blocks <- list(...)
    blocks <- unlist(blocks, use.names = FALSE)
    paste(blocks, collapse = "")
  }

  calculate_aic_bic <- function(model, model_name = "") {
    if (is.null(model) || inherits(model, "coeftest")) return(c(AIC = NA, BIC = NA))
    if (inherits(model, "fixest")) return(c(AIC = tryCatch(AIC(model), error = function(e) NA), BIC = tryCatch(BIC(model), error = function(e) NA)))
    if (inherits(model, "plm")) {
      tryCatch({
          res <- residuals(model)
          n <- length(res)
          sigma2 <- sum(res^2) / n
          loglik <- -n / 2 * log(2 * pi) - n / 2 * log(sigma2) - n / 2
          k_slope <- length(coef(model))
          eff <- tryCatch(model$args$effect, error = function(e) {
            tryCatch(attr(model, "effect"), error = function(e2) "individual")
          })
          mod <- tryCatch(model$args$model, error = function(e) "within")
          k_fe <- 0
          if (mod == "within") {
            n_id <- length(unique(model$model[, 1]))
            n_ti <- length(unique(model$model[, 2]))
            if (eff == "individual") {
              k_fe <- n_individuals
            } else if (eff == "time") {
              k_fe <- n_periods
            } else if (eff == "twoways") k_fe <- n_individuals + n_periods - 1
          }
          k <- k_slope + k_fe
          aic <- -2 * loglik + 2 * k
          bic <- -2 * loglik + k * log(n)
          c(AIC = aic, BIC = bic)
        },
        error = function(e) c(AIC = NA, BIC = NA)
      )
    } else if (inherits(model, "lm")) {
      c(AIC = tryCatch(AIC(model), error = function(e) NA), BIC = tryCatch(BIC(model), error = function(e) NA))
    } else {
      c(AIC = NA, BIC = NA)
    }
  }
  aic_bic_results <- data.frame(Model = character(), AIC = numeric(), BIC = numeric(), stringsAsFactors = FALSE)
  
  # 读取数据
  mydata <- read.xlsx(input_xlsx_path)
  if (ncol(mydata) < 4) stop("数据不符合要求，至少需要4列：第1列id，第2列time，第3列Y，第4列及以后为X。")
  myColNames <- colnames(mydata)
  newColTTTNames <- gsub("\\W", "", myColNames)
  newColNames <- gsub("^", "v_", newColTTTNames)
  colnames(mydata) <- newColNames
  myColNames <- colnames(mydata)

  index_id <- myColNames[1]
  index_time <- myColNames[2]
  model_Y <- myColNames[3]
  varX_vectors <- myColNames[!myColNames %in% c(index_id, index_time, model_Y)]
  if (length(varX_vectors) < 1) stop("数据不符合要求，至少需要4列：第1列id，第2列time，第3列Y，第4列及以后为X。")

  na_count <- sum(is.na(mydata[, c(model_Y, varX_vectors)]))  
  # 简化处理缺失值：用0填充（注意：可能导致估计偏误）
  mydata[is.na(mydata)] <- 0

  model_X <- paste(varX_vectors, collapse = " + ")
  str_reg_formula <- paste(model_Y, "~", model_X)
  model_formula <- as.formula(str_reg_formula)

  pdata <- tryCatch(
    pdata.frame(mydata, index = c(index_id, index_time)),
    error = function(e) NULL
  )

  n_individuals <- length(unique(mydata[[index_id]]))
  n_periods <- length(unique(mydata[[index_time]]))
  n_total <- nrow(mydata)
  is_balanced <- (n_total == n_individuals * n_periods)

  panel_info <- data.frame(
    项目 = c("个体数(N)", "时间期数(T)", "总观测数(N*T)", "是否平衡面板", "个体变量", "时间变量", "因变量", "解释变量"),
    内容 = c(n_individuals, n_periods, n_total, ifelse(is_balanced, "是", "否（非平衡面板）"), index_id, index_time, model_Y, paste(varX_vectors, collapse = ", ")),
    stringsAsFactors = FALSE
  )
  
  # 描述统计
  col_to_drops <- c(index_id, index_time)
  describe_data <- mydata[, !(names(mydata) %in% col_to_drops), drop = FALSE]
  desTable <- descr(describe_data, headings = TRUE, stats = c("mean", "med", "sd", "min", "max", "n.valid", "pct.valid"), transpose = TRUE)
  desTable_re_index <- cbind(rownames(desTable), data.frame(desTable, row.names = NULL))
  colnames(desTable_re_index)[1] <- "Variable"
  cn_names <- c("Variable" = "变量", "Mean" = "平均值", "Median" = "中位数", "Std.Dev" = "标准差", "Min" = "最小值", "Max" = "最大值", "N.Valid" = "有效观测值", "Pct.Valid" = "有效观测值百分比")
  for (old_name in names(cn_names)) {
    if (old_name %in% colnames(desTable_re_index)) {
      colnames(desTable_re_index)[colnames(desTable_re_index) == old_name] <- cn_names[old_name]
    }
  }

  # 相关系数矩阵
  cor_data <- describe_data
  for (nm in names(cor_data)) cor_data[[nm]] <- as.numeric(cor_data[[nm]])
  cor_matrix <- cor(cor_data, use = "pairwise.complete.obs")
  cor_table <- as.data.frame(cor_matrix)
  cor_table <- cbind(变量 = rownames(cor_table), cor_table)
  rownames(cor_table) <- NULL

  # OLS（基准）
  ols_model <- lm(model_formula, data = mydata)
  ols_ab <- calculate_aic_bic(ols_model)
  aic_bic_results <- rbind(aic_bic_results, data.frame(Model = "OLS", AIC = ols_ab["AIC"], BIC = ols_ab["BIC"]))

  # 面板模型
  end_list <- list()
  end_vect <- c()

  # Pooled
  pool_model <- safe_fit(quote(
    plm(model_formula, data = mydata, index = c(index_id, index_time), model = "pooling")
  ))
  if (!is.null(pool_model)) {
    end_list[[length(end_list) + 1]] <- pool_model
    end_vect <- append(end_vect, "Pooled")
    ab <- calculate_aic_bic(pool_model)
    aic_bic_results <- rbind(aic_bic_results, data.frame(Model = "Pooled", AIC = ab["AIC"], BIC = ab["BIC"]))
  }

  # Individual FE
  fixed_model_ObsEff <- safe_fit(quote(
    plm(model_formula, data = mydata, index = c(index_id, index_time), model = "within", effect = "individual")
  ))
  if (!is.null(fixed_model_ObsEff)) {
    end_list[[length(end_list) + 1]] <- fixed_model_ObsEff
    end_vect <- append(end_vect, "Individual FE")
    ab <- calculate_aic_bic(fixed_model_ObsEff)
    aic_bic_results <- rbind(aic_bic_results, data.frame(Model = "Individual FE", AIC = ab["AIC"], BIC = ab["BIC"]))
  }

  # Time FE
  fixed_model_timeEff <- safe_fit(quote(
    plm(model_formula, data = mydata, index = c(index_id, index_time), model = "within", effect = "time")
  ))
  if (!is.null(fixed_model_timeEff)) {
    end_list[[length(end_list) + 1]] <- fixed_model_timeEff
    end_vect <- append(end_vect, "Time FE")
    ab <- calculate_aic_bic(fixed_model_timeEff)
    aic_bic_results <- rbind(aic_bic_results, data.frame(Model = "Time FE", AIC = ab["AIC"], BIC = ab["BIC"]))
  }

  # Twoway FE
  fixed_model_twoEff <- safe_fit(quote(
    plm(model_formula, data = mydata, index = c(index_id, index_time), model = "within", effect = "twoways")
  ))
  if (!is.null(fixed_model_twoEff)) {
    end_list[[length(end_list) + 1]] <- fixed_model_twoEff
    end_vect <- append(end_vect, "Twoway FE")
    ab <- calculate_aic_bic(fixed_model_twoEff)
    aic_bic_results <- rbind(aic_bic_results, data.frame(Model = "Twoway FE", AIC = ab["AIC"], BIC = ab["BIC"]))
  }

  # 固定效应选择
  F_obsEff_pval <- NA_real_
  F_timeEff_pval <- NA_real_
  F_twoEff_pval <- NA_real_
  if (!is.null(fixed_model_ObsEff) && !is.null(pool_model)) {
    F_obsEff_pval <- tryCatch(
      as.numeric(pFtest(fixed_model_ObsEff, pool_model)$p.value),
      error = function(e) NA_real_
    )
  }
  if (!is.null(fixed_model_timeEff) && !is.null(pool_model)) {
    F_timeEff_pval <- tryCatch(
      as.numeric(pFtest(fixed_model_timeEff, pool_model)$p.value),
      error = function(e) NA_real_
    )
  }
  if (!is.null(fixed_model_twoEff) && !is.null(pool_model)) {
    F_twoEff_pval <- tryCatch(
      as.numeric(pFtest(fixed_model_twoEff, pool_model)$p.value),
      error = function(e) NA_real_
    )
  }

  # 选择固定效应类型
  fix_effect_choice <- "individual"
  if (!is.na(F_obsEff_pval) && !is.na(F_timeEff_pval) &&
    F_obsEff_pval < 0.05 && F_timeEff_pval < 0.05) {
    fix_effect_choice <- "twoways"
  } else if (!is.na(F_timeEff_pval) && F_timeEff_pval < 0.05 &&
    (is.na(F_obsEff_pval) || F_obsEff_pval >= 0.05)) {
    fix_effect_choice <- "time"
  } else if (!is.na(F_obsEff_pval) && F_obsEff_pval < 0.05) {
    fix_effect_choice <- "individual"
  }

  fe_selection_tests <- data.frame(
    检验 = c("个体FE联合显著性(F)", "时间FE联合显著性(F)", "双向FE联合显著性(F)"),
    p_value = c(F_obsEff_pval, F_timeEff_pval, F_twoEff_pval),
    通过 = c(pass_if_p_less_005(F_obsEff_pval), pass_if_p_less_005(F_timeEff_pval), pass_if_p_less_005(F_twoEff_pval)),
    说明 = c("p<0.05则个体FE显著", "p<0.05则时间FE显著", "p<0.05则双向FE显著"),
    stringsAsFactors = FALSE
  )
  fe_selection_tests$结论 <- judge_text(fe_selection_tests$通过)

  fixed_model <- switch(fix_effect_choice, "individual" = fixed_model_ObsEff, "time" = fixed_model_timeEff, "twoways" = fixed_model_twoEff)

  if (!is.null(fixed_model) && !(fix_effect_choice %in% c("individual", "time", "twoways")[
    c(!is.null(fixed_model_ObsEff), !is.null(fixed_model_timeEff), !is.null(fixed_model_twoEff))
  ])) {    
    fixed_model <- safe_fit(quote(plm(model_formula, data = mydata, index = c(index_id, index_time), model = "within", effect = fix_effect_choice)))
  }
  if (!is.null(fixed_model)) {
    end_list[[length(end_list) + 1]] <- fixed_model
    end_vect <- append(end_vect, paste0("Selected FE: ", fix_effect_choice))
  }

  # 随机效应
  random_model <- NULL
  for (rm_method in c("swar", "amemiya", "walhus", "nerlove")) {
    random_model <- tryCatch(
      plm(model_formula, data = mydata, index = c(index_id, index_time), model = "random", effect = fix_effect_choice, random.method = rm_method),
      error = function(e) NULL
    )
    if (!is.null(random_model)) break
  }

  if (is.null(random_model) && fix_effect_choice == "twoways") {
    for (rm_method in c("swar", "amemiya", "walhus", "nerlove")) {
      random_model <- tryCatch(
        plm(model_formula, data = mydata, index = c(index_id, index_time), model = "random", effect = "individual", random.method = rm_method),
        error = function(e) NULL
      )
      if (!is.null(random_model)) break
    }
  }

  if (!is.null(random_model)) {
    end_list[[length(end_list) + 1]] <- random_model
    end_vect <- append(end_vect, "Random Effects")
    ab <- calculate_aic_bic(random_model)
    aic_bic_results <- rbind(aic_bic_results, data.frame(Model = "Random Effects", AIC = ab["AIC"], BIC = ab["BIC"]))
  }

  # 模型选择检验
  model_F_test_pval <- NA_real_
  model_hausman_test_pval <- NA_real_
  model_LM_test_pval <- NA_real_

  if (!is.null(fixed_model) && !is.null(pool_model)) {
    model_F_test_pval <- tryCatch(
      as.numeric(pFtest(fixed_model, pool_model)$p.value),
      error = function(e) NA_real_
    )
  }

  if (!is.null(fixed_model) && !is.null(random_model)) {
    model_hausman_test_pval <- tryCatch(
      as.numeric(phtest(fixed_model, random_model)$p.value),
      error = function(e) {        
        tryCatch({
            as.numeric(phtest(fixed_model, random_model, method = "aux")$p.value)
          },
          error = function(e2) NA_real_
        )
      }
    )
  }

  if (!is.null(pool_model)) {
    model_LM_test_pval <- tryCatch(
      as.numeric(plmtest(pool_model, type = "bp")$p.value),
      error = function(e) NA_real_
    )
  }

  # 模型选择结论
  selected_model_type <- "Pooled OLS"
  if (!is.na(model_F_test_pval) && model_F_test_pval < 0.05) {
    selected_model_type <- "Fixed Effects"
    if (!is.na(model_hausman_test_pval) && model_hausman_test_pval >= 0.05) {
      selected_model_type <- "Random Effects"
    }
  } else if (!is.na(model_LM_test_pval) && model_LM_test_pval < 0.05) {
    selected_model_type <- "Random Effects"
  }

  model_select_tests <- data.frame(
    检验 = c("F检验(Pooled vs FE)", "Hausman检验(FE vs RE)", "BP-LM检验(Pooled vs RE)", "选择结论"),
    p_value = c(model_F_test_pval, model_hausman_test_pval, model_LM_test_pval, NA),
    说明 = c("p<0.05拒绝Pooled，选择FE", "p<0.05拒绝RE，选择FE；p>=0.05选择RE", "p<0.05拒绝Pooled，选择RE", selected_model_type),
    stringsAsFactors = FALSE
  )

  # 一阶差分
  first_diff_model <- safe_fit(quote(plm(model_formula, data = mydata, index = c(index_id, index_time), model = "fd")))
  if (!is.null(first_diff_model)) {
    end_list[[length(end_list) + 1]] <- first_diff_model
    end_vect <- append(end_vect, "First-Difference")
    ab <- calculate_aic_bic(first_diff_model)
    aic_bic_results <- rbind(aic_bic_results, data.frame(Model = "First-Difference", AIC = ab["AIC"], BIC = ab["BIC"]))
  }

  # 截面相关
  assump_csd_cd_pval <- tryCatch({
      test_model <- if (!is.null(fixed_model_twoEff)) {
        fixed_model_twoEff
      } else if (!is.null(fixed_model_ObsEff)) {
        fixed_model_ObsEff
      } else {
        NULL
      }
      if (is.null(test_model)) {
        NA_real_
      } else {
        as.numeric(pcdtest(test_model, test = "cd")$p.value)
      }
    },
    error = function(e) NA_real_
  )

  assump_csd_lm_pval <- tryCatch({
      test_model <- if (!is.null(fixed_model_twoEff)) {
        fixed_model_twoEff
      } else if (!is.null(fixed_model_ObsEff)) {
        fixed_model_ObsEff
      } else {
        NULL
      }
      if (is.null(test_model)) {
        NA_real_
      } else {
        as.numeric(pcdtest(test_model, test = "lm")$p.value)
      }
    },
    error = function(e) NA_real_
  )

  # 序列相关
  assump_sc_fe_pval <- tryCatch({
      if (is.null(fixed_model)) {
        NA_real_
      } else {
        as.numeric(pbgtest(fixed_model)$p.value)
      }
    },
    error = function(e) NA_real_
  )

  assump_sc_pool_pval <- tryCatch({
      if (is.null(pool_model)) {
        NA_real_
      } else {
        as.numeric(pbgtest(pool_model)$p.value)
      }
    },
    error = function(e) NA_real_
  )

  assump_sc_re_pval <- tryCatch({
      if (is.null(random_model)) {
        NA_real_
      } else {
        as.numeric(pbgtest(random_model)$p.value)
      }
    },
    error = function(e) NA_real_
  )

  assump_wooldridge_pval <- tryCatch({
      if (is.null(fixed_model)) {
        NA_real_
      } else {
        as.numeric(pwartest(fixed_model)$p.value)
      }
    },
    error = function(e) NA_real_
  )

  # 单位根检验
  if (is.null(pdata)) {
    pdata <- tryCatch(
      pdata.frame(mydata, index = c(index_id, index_time)),
      error = function(e) NULL
    )
  }

  all_vars_for_ur <- c(model_Y, varX_vectors)
  unit_root_results <- data.frame(变量 = character(), 检验方法 = character(), p_value = numeric(), 结论 = character(), stringsAsFactors = FALSE)

  if (!is.null(pdata)) {
    for (var_name in all_vars_for_ur) {
      var_series <- tryCatch(pdata[[var_name]], error = function(e) NULL)
      if (is.null(var_series)) next

      adf_p <- tryCatch({
          as.numeric(tseries::adf.test(as.numeric(var_series), k = 2)$p.value)
        },
        error = function(e) NA_real_
      )
      
      mw_p <- tryCatch({
          as.numeric(purtest(var_series, pmax = 4, exo = "intercept", test = "madwu")$statistic$p.value)
        },
        error = function(e) NA_real_
      )

      ips_p <- tryCatch({
          as.numeric(purtest(var_series, pmax = 4, exo = "intercept", test = "ips")$statistic$p.value)
        },
        error = function(e) NA_real_
      )

      llc_p <- tryCatch({
          as.numeric(purtest(var_series, pmax = 4, exo = "intercept", test = "levinlin")$statistic$p.value)
        },
        error = function(e) NA_real_
      )

      hadri_p <- tryCatch({
          as.numeric(purtest(var_series, exo = "intercept", test = "hadri")$statistic$p.value)
        },
        error = function(e) NA_real_
      )

      unit_root_results <- rbind(unit_root_results, data.frame(
        变量 = rep(var_name, 5),
        检验方法 = c("ADF", "Maddala-Wu", "IPS", "Levin-Lin-Chu", "Hadri"),
        p_value = c(adf_p, mw_p, ips_p, llc_p, hadri_p),
        原假设 = c("有单位根", "有单位根", "有单位根", "有单位根", "无单位根(平稳)"),
        通过 = c(
          pass_if_p_less_005(adf_p),
          pass_if_p_less_005(mw_p),
          pass_if_p_less_005(ips_p),
          pass_if_p_less_005(llc_p),
          pass_if_p_greater_005(hadri_p)
        ),
        stringsAsFactors = FALSE
      ))
    }
    unit_root_results$结论 <- judge_text(unit_root_results$通过)
  }

  # 异方差检验
  assump_hete_bp_pval <- tryCatch({
      if (is.null(pool_model)) {
        NA_real_
      } else {
        as.numeric(bptest(model_formula, data = mydata, studentize = TRUE)$p.value)
      }
    },
    error = function(e) NA_real_
  )

  assump_hete_groupwise_pval <- tryCatch({
      if (is.null(fixed_model_ObsEff)) {
        NA_real_
      } else {
        fe_resid <- residuals(fixed_model_ObsEff)
        idx <- index(fixed_model_ObsEff)
        id_index <- idx[[1]]
        if (is.null(id_index) || length(unique(id_index)) < 2) {
          NA_real_
        } else {
          as.numeric(bartlett.test(as.numeric(fe_resid) ~ factor(id_index))$p.value)
        }
      }
    },
    error = function(e) NA_real_
  )

  # 假设检验汇总表
  assumptions_tests <- data.frame(
    类别 = c("截面相关", "截面相关", "序列相关", "序列相关", "序列相关", "序列相关", "异方差", "异方差"),
    检验方法 = c("Pesaran CD test", "Breusch-Pagan LM test", "BG/Wooldridge(FE)", "BG/Wooldridge(Pooled)", "BG/Wooldridge(RE)", "Wooldridge FE AR(1)", "Breusch-Pagan(Pooled)", "Bartlett组间异方差"),
    p_value = c(assump_csd_cd_pval, assump_csd_lm_pval, assump_sc_fe_pval, assump_sc_pool_pval, assump_sc_re_pval, assump_wooldridge_pval, assump_hete_bp_pval, assump_hete_groupwise_pval),
    原假设 = c("无截面相关", "无截面相关", "无序列相关", "无序列相关", "无序列相关", "无序列相关", "同方差", "各组方差相同"),
    通过 = c(
      pass_if_p_greater_005(assump_csd_cd_pval),
      pass_if_p_greater_005(assump_csd_lm_pval),
      pass_if_p_greater_005(assump_sc_fe_pval),
      pass_if_p_greater_005(assump_sc_pool_pval),
      pass_if_p_greater_005(assump_sc_re_pval),
      pass_if_p_greater_005(assump_wooldridge_pval),
      pass_if_p_greater_005(assump_hete_bp_pval),
      pass_if_p_greater_005(assump_hete_groupwise_pval)
    ),
    stringsAsFactors = FALSE
  )
  assumptions_tests$结论 <- judge_text(assumptions_tests$通过)

  # 稳健标准误
  # 固定效应 + 稳健标准误
  fixed_model_hac <- tryCatch({
      if (is.null(fixed_model)) {
        NULL
      } else {
        coeftest(fixed_model, vcovHC(fixed_model, method = "arellano", type = "HC1"))
      }
    },
    error = function(e) NULL
  )

  fixed_model_cluster_id <- tryCatch({
      if (is.null(fixed_model)) {
        NULL
      } else {
        coeftest(fixed_model, vcovHC(fixed_model, type = "HC1", cluster = "group"))
      }
    },
    error = function(e) NULL
  )

  fixed_model_cluster_time <- tryCatch({
      if (is.null(fixed_model)) {
        NULL
      } else {
        coeftest(fixed_model, vcovHC(fixed_model, type = "HC1", cluster = "time"))
      }
    },
    error = function(e) NULL
  )

  fixed_model_pcse <- tryCatch({
      if (is.null(fixed_model)) {
        NULL
      } else {
        coeftest(fixed_model, vcovBK(fixed_model, type = "HC3", cluster = "group"))
      }
    },
    error = function(e) NULL
  )

  fixed_model_scc <- tryCatch({
      if (is.null(fixed_model)) {
        NULL
      } else {
        coeftest(fixed_model, vcovSCC(fixed_model, type = "HC3"))
      }
    },
    error = function(e) NULL
  )

  fixed_model_vcovDC <- tryCatch({
      if (is.null(fixed_model)) {
        NULL
      } else {
        coeftest(fixed_model, vcovDC(fixed_model, type = "HC3"))
      }
    },
    error = function(e) NULL
  )

  # 随机效应 + 稳健标准误
  random_model_hc <- tryCatch({
      if (is.null(random_model)) {
        NULL
      } else {
        coeftest(random_model, vcovHC(random_model, type = "HC1"))
      }
    },
    error = function(e) NULL
  )

  # fixest 双向固定效应 + 双重聚类
  str_fixest_formula <- paste(str_reg_formula, "|", index_id, "+", index_time)
  fixest_formula <- as.formula(str_fixest_formula)
  fixest_model <- safe_fit(quote(feols(fixest_formula, data = mydata, cluster = c(index_id, index_time))))

  if (!is.null(fixest_model)) {
    end_list[[length(end_list) + 1]] <- fixest_model
    end_vect <- append(end_vect, "Twoway FE (fixest, double cluster)")
    ab <- calculate_aic_bic(fixest_model)
    aic_bic_results <- rbind(aic_bic_results, data.frame(Model = "Twoway FE (fixest)", AIC = ab["AIC"], BIC = ab["BIC"]))
  }

  # 结果汇总表
  get_panel_model_stats <- function(model, model_name) {
    null_result <- list(
      model = model_name,
      coef = setNames(numeric(0), character(0)),
      se = setNames(numeric(0), character(0)),
      nobs = NA_real_
    )
    if (is.null(model)) return(null_result)

    if (inherits(model, "coeftest")) {
      cf <- tryCatch(as.matrix(model), error = function(e) NULL)
      if (is.null(cf) || nrow(cf) == 0) {
        return(null_result)
      }
      return(list(
        model = model_name,
        coef = setNames(as.numeric(cf[, 1]), rownames(cf)),
        se = setNames(as.numeric(cf[, 2]), rownames(cf)),
        nobs = NA_real_
      ))
    }

    if (inherits(model, "fixest")) {
      sm <- tryCatch(summary(model), error = function(e) NULL)
      if (is.null(sm)) return(null_result)
      cf <- sm$coeftable
      return(list(
        model = model_name,
        coef = setNames(as.numeric(cf[, 1]), rownames(cf)),
        se = setNames(as.numeric(cf[, 2]), rownames(cf)),
        nobs = tryCatch(sm$nobs, error = function(e) NA_real_)
      ))
    }

    sm <- tryCatch(summary(model), error = function(e) NULL)
    if (is.null(sm) || is.null(sm$coefficients)) return(null_result)
    cf <- sm$coefficients
    if (is.vector(cf)) cf <- matrix(cf, nrow = 1)
    list(
      model = model_name,
      coef = setNames(as.numeric(cf[, 1]), rownames(cf)),
      se = setNames(if (ncol(cf) >= 2) as.numeric(cf[, 2]) else rep(NA_real_, nrow(cf)), rownames(cf)),
      nobs = tryCatch(nobs(model), error = function(e) {
        tryCatch(length(residuals(model)), error = function(e2) NA_real_)
      })
    )
  }

  get_panel_p_named <- function(model) {
    if (is.null(model)) return(setNames(numeric(0), character(0)))
    if (inherits(model, "coeftest")) {
      cf <- tryCatch(as.matrix(model), error = function(e) NULL)
      if (is.null(cf) || nrow(cf) == 0) return(setNames(numeric(0), character(0)))
      if (ncol(cf) >= 4) return(setNames(as.numeric(cf[, 4]), rownames(cf)))
      return(setNames(rep(NA_real_, nrow(cf)), rownames(cf)))
    }
    if (inherits(model, "fixest")) {
      sm <- tryCatch(summary(model), error = function(e) NULL)
      if (is.null(sm)) return(setNames(numeric(0), character(0)))
      cf <- sm$coeftable
      if (ncol(cf) >= 4) return(setNames(as.numeric(cf[, 4]), rownames(cf)))
      return(setNames(rep(NA_real_, nrow(cf)), rownames(cf)))
    }

    sm <- tryCatch(summary(model), error = function(e) NULL)
    if (is.null(sm) || is.null(sm$coefficients)) return(setNames(numeric(0), character(0)))
    cf <- sm$coefficients
    if (is.vector(cf)) cf <- matrix(cf, nrow = 1)
    if (ncol(cf) >= 4) return(setNames(as.numeric(cf[, 4]), rownames(cf)))
    return(setNames(rep(NA_real_, nrow(cf)), rownames(cf)))
  }

  paperStyle_models <- list(
    get_panel_model_stats(pool_model, "Pooled Model(混合面板模型)"),
    get_panel_model_stats(fixed_model_ObsEff, "Individual FE(个体固定效应模型)"),
    get_panel_model_stats(fixed_model_timeEff, "Time FE(时间固定效应模型)"),
    get_panel_model_stats(fixed_model_twoEff, "Two-way FE(双向固定效应模型)"),
    get_panel_model_stats(random_model, "Random Effects(随机效应模型)"),
    get_panel_model_stats(first_diff_model, "First-Difference(一阶差分模型)"),
    get_panel_model_stats(fixed_model_hac, "FE + Arellano HAC(固定效应 + 异方差序列相关稳健标准误)"),
    get_panel_model_stats(fixed_model_cluster_id, "FE + ClusterID(固定效应 + 个体聚类稳健标准误)"),
    get_panel_model_stats(fixed_model_cluster_time, "FE + ClusterTime(固定效应 + 时间聚类稳健标准误)"),
    get_panel_model_stats(fixed_model_pcse, "FE + PCSE(固定效应 + 面板校正标准误(Beck-Katz))"),
    get_panel_model_stats(fixed_model_scc, "FE + Driscoll-Kraay(固定效应 + 空间相关稳健标准误(SCC))"),
    get_panel_model_stats(fixed_model_vcovDC, "FE + Double Cluster(固定效应 + 双重聚类稳健标准误)"),
    get_panel_model_stats(random_model_hc, "RE + Robust SE(随机效应 + 稳健标准误)"),
    get_panel_model_stats(fixest_model, "Fixest Twoway FE(使用Fixest包 + 双向固定效应 + 双重聚类")
  )

  paperStyle_pvals <- list(
    get_panel_p_named(pool_model),
    get_panel_p_named(fixed_model_ObsEff),
    get_panel_p_named(fixed_model_timeEff),
    get_panel_p_named(fixed_model_twoEff),
    get_panel_p_named(random_model),
    get_panel_p_named(first_diff_model),
    get_panel_p_named(fixed_model_hac),
    get_panel_p_named(fixed_model_cluster_id),
    get_panel_p_named(fixed_model_cluster_time),
    get_panel_p_named(fixed_model_pcse),
    get_panel_p_named(fixed_model_scc),
    get_panel_p_named(fixed_model_vcovDC),
    get_panel_p_named(random_model_hc),
    get_panel_p_named(fixest_model)
  )

  all_coef_names <- unique(unlist(lapply(paperStyle_models, function(m) names(m$coef))))
  if (length(all_coef_names) == 0) all_coef_names <- "(Intercept)"

  paperStyle_col_names <- c("项", sapply(paperStyle_models, function(m) m$model))
  paperStyle_table <- data.frame(matrix(ncol = length(paperStyle_col_names), nrow = 0), stringsAsFactors = FALSE)
  names(paperStyle_table) <- paperStyle_col_names

  for (coef_name in all_coef_names) {
    coef_row <- c(coef_name)
    se_row <- c(" ")
    for (j in seq_along(paperStyle_models)) {
      m <- paperStyle_models[[j]]
      pvec <- paperStyle_pvals[[j]]
      est <- if (coef_name %in% names(m$coef)) m$coef[[coef_name]] else NA_real_
      se0 <- if (coef_name %in% names(m$se)) m$se[[coef_name]] else NA_real_
      pv <- if (coef_name %in% names(pvec)) pvec[[coef_name]] else NA_real_
      coef_row <- c(coef_row, ifelse(is.na(est), "", paste0(format(round(est, 4), nsmall = 4, trim = TRUE), make_sig_star(pv))))
      se_row <- c(se_row, ifelse(is.na(se0), "", paste0("(", format(round(se0, 4), nsmall = 4, trim = TRUE), ")")))
    }
    paperStyle_table <- rbind(paperStyle_table, make_paperStyle_row(coef_row, paperStyle_col_names), make_paperStyle_row(se_row, paperStyle_col_names))
  }

  extra_rows <- list(c("Num. obs.", sapply(paperStyle_models, function(m) fmt0(m$nobs))), c("说明", rep("", length(paperStyle_models))))
  for (rw in extra_rows) paperStyle_table <- rbind(paperStyle_table, make_paperStyle_row(rw, paperStyle_col_names))
  paperStyle_table[nrow(paperStyle_table), 1] <- "Standard errors in parentheses. *** p<0.01; ** p<0.05; * p<0.1."

  # 结果提示
  has_csd <- (!is.na(assump_csd_cd_pval) && assump_csd_cd_pval < 0.05)
  has_sc <- (!is.na(assump_sc_fe_pval) && assump_sc_fe_pval < 0.05)
  has_het <- (!is.na(assump_hete_bp_pval) && assump_hete_bp_pval < 0.05)

  robust_recommendation <- "标准OLS/FE标准误"
  if (has_csd && has_sc && has_het) {
    robust_recommendation <- "推荐Driscoll-Kraay SCC或双重聚类稳健标准误"
  } else if (has_csd && has_het) {
    robust_recommendation <- "推荐PCSE(Beck-Katz)或Driscoll-Kraay SCC"
  } else if (has_sc && has_het) {
    robust_recommendation <- "推荐Arellano HAC稳健标准误"
  } else if (has_het) {
    robust_recommendation <- "推荐按个体聚类的HC稳健标准误"
  } else if (has_sc) {
    robust_recommendation <- "推荐Newey-West HAC或Arellano稳健标准误"
  } else if (has_csd) {
    robust_recommendation <- "推荐PCSE或Driscoll-Kraay SCC"
  }

  final_hint <- data.frame(
    项目 = c("模型选择结论", "固定效应类型", "稳健标准误建议", "截面相关", "序列相关", "异方差", "面板结构", "单位根提示"),
    内容 = c(
      selected_model_type,
      fix_effect_choice,
      robust_recommendation,
      ifelse(has_csd, "存在截面相关，需使用PCSE/SCC", "未检测到显著截面相关"),
      ifelse(has_sc, "存在序列相关，需使用HAC/SCC", "未检测到显著序列相关"),
      ifelse(has_het, "存在异方差，需使用稳健标准误", "未检测到显著异方差"),
      ifelse(is_balanced, "平衡面板", "非平衡面板，部分检验可能不适用"),
      ifelse(n_individuals < 10, "N较小，IPS/LLC等大N检验结果仅供参考", "")
    ),
    stringsAsFactors = FALSE
  )

  # Excel输出
  xls_output_list_in_Fun <- list(
    "数据" = mydata,
    "面板结构" = panel_info,
    "描述统计" = desTable_re_index,
    "相关系数矩阵" = cor_table,
    "固定效应选择" = fe_selection_tests,
    "模型选择" = model_select_tests,
    "单位根检验" = unit_root_results,
    "假设检验" = assumptions_tests,
    "结果汇总" = paperStyle_table,
    "模型比较" = aic_bic_results,
    "结果提示" = final_hint
  )

  # html输出
  logBlocksHtml <- character(0)

  logBlocksHtml <- c(
    logBlocksHtml,
    info_table_to_html_block(
      list(
        "分析类型" = "面板数据分析",
        "个体变量" = index_id, "时间变量" = index_time,
        "因变量" = model_Y,
        "解释变量" = paste(varX_vectors, collapse = ", "),
        "总观测数" = n_total,
        "个体数" = n_individuals, "时间期数" = n_periods,
        "平衡面板" = ifelse(is_balanced, "是", "否")
      ),
      title = "分析说明",
      desc = NULL
    )
  )

  logBlocksHtml <- c(
    logBlocksHtml,
    df_to_html_block(desTable_re_index, "描述统计表", NULL)
  )

  logBlocksHtml <- c(
    logBlocksHtml,
    df_to_html_block(cor_table, "相关系数矩阵", "Pearson相关系数")
  )

  if (!is.null(ols_model)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("OLS(基准，未考虑面板结构)", paste0(str_reg_formula), summary(ols_model))
    )
  }

  if (!is.null(pool_model)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("混合模型", NULL, summary(pool_model))
    )
  }

  if (!is.null(fixed_model_ObsEff)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("个体固定效应模型", NULL, summary(fixed_model_ObsEff))
    )
  }

  if (!is.null(fixed_model_timeEff)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("时间固定效应模型", NULL, summary(fixed_model_timeEff))
    )
  }

  if (!is.null(fixed_model_twoEff)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("双向固定效应模型", NULL, summary(fixed_model_twoEff))
    )
  }

  logBlocksHtml <- c(
    logBlocksHtml,
    df_to_html_block(fe_selection_tests, "固定效应类型选择", paste0("基于F检验选择 ", fix_effect_choice))
  )

  if (!is.null(random_model)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block(
        "随机效应模型",
        paste0("effect='", fix_effect_choice, "'，与固定效应模型的effect类型一致"),
        summary(random_model)
      )
    )
  }

  logBlocksHtml <- c(
    logBlocksHtml,
    df_to_html_block(model_select_tests, "模型选择检验", NULL)
  )

  if (!is.null(first_diff_model)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("一阶差分模型", "剔除个体固定效应", summary(first_diff_model))
    )
  }

  logBlocksHtml <- c(
    logBlocksHtml,
    df_to_html_block(
      unit_root_results, "单位根检验(所有变量)",
      "ADF/MW/IPS: H0=有单位根，p<0.05拒绝即平稳；Hadri: H0=平稳，p<0.05拒绝即非平稳"
    )
  )

  logBlocksHtml <- c(
    logBlocksHtml,
    df_to_html_block(assumptions_tests, "假设检验汇总表", NULL)
  )

  if (!is.null(fixed_model_hac)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("FE + Arellano HAC稳健标准误", "对异方差和序列相关稳健", fixed_model_hac)
    )
  }

  if (!is.null(fixed_model_cluster_id)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("FE + 按个体聚类稳健标准误", "HC1, cluster='group'", fixed_model_cluster_id)
    )
  }

  if (!is.null(fixed_model_cluster_time)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("FE + 按时间聚类稳健标准误", "HC1, cluster='time'", fixed_model_cluster_time)
    )
  }

  if (!is.null(fixed_model_pcse)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("FE + PCSE(Beck-Katz)", "对截面相关和异方差稳健", fixed_model_pcse)
    )
  }

  if (!is.null(fixed_model_scc)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("FE + Driscoll-Kraay SCC", "对截面相关、序列相关与异方差同时稳健", fixed_model_scc)
    )
  }

  if (!is.null(fixed_model_vcovDC)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("FE + 双重聚类稳健标准误", "同时按个体和时间聚类", fixed_model_vcovDC)
    )
  }

  if (!is.null(random_model_hc)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block("RE + HC稳健标准误", "对异方差稳健", random_model_hc)
    )
  }

  if (!is.null(fixest_model)) {
    logBlocksHtml <- c(
      logBlocksHtml,
      text_log_block(
        "使用Fixest包估计的双向固定效应 + 双重聚类",
        paste0("公式 ", str_fixest_formula),
        summary(fixest_model)
      )
    )
  }

  logBlocksHtml <- c(
    logBlocksHtml,
    df_to_html_block(paperStyle_table, "结果汇总表", NULL)
  )

  logBlocksHtml <- c(
    logBlocksHtml,
    df_to_html_block(aic_bic_results, "模型比较(AIC/BIC)", NULL)
  )

  logBlocksHtml <- c(
    logBlocksHtml,
    df_to_html_block(final_hint, "结果提示与建议", NULL)
  )

  logHtml_in_Fun <- combine_log_blocks(logBlocksHtml)
  
  return(list(xls_output_list = xls_output_list_in_Fun, logHtml = logHtml_in_Fun))
}

# 使用说明
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

xlsxFilesNameStr <- xlsxFilesVec[1]
outPutFileNameXlsx <- sub("\\.xlsx$", "_OutputGit.xlsx", xlsxFilesNameStr)
outPutFileNameHtml <- sub("\\.xlsx$", "_OutputGit.html", xlsxFilesNameStr)

# 运行分析
funAll_return_REsult <- myPanelDataAnalysis_Function(xlsxFilesNameStr)

# 保存 Excel
write.xlsx(
  funAll_return_REsult$xls_output_list,
  file = outPutFileNameXlsx,
  startCol = 1, startRow = 1,
  rowNames = FALSE, asTable = FALSE, overwrite = TRUE
)

# 保存 HTML
full_html <- paste0(
  '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>分析日志</title>',
  "<style>",
  "body{margin:0;padding:0;font-family:monospace;background:#f0f0f0;}",
  ".log-container{margin:15px;background:#f5f5f5;border-radius:4px;padding:10px;",
  "  max-width:calc(100vw - 30px);",
  "  overflow-x:auto;}",
  ".log-title{font-size:13px;margin:10px 0 4px 0;}",
  ".log-desc{font-size:12px;margin:0 0 4px 0;}",
  ".log-block{margin-bottom:12px;}",
  ".log-text{white-space:pre;font-size:12px;margin:0;}",
  ".log-table-block{margin-bottom:12px;}",
  ".log-table-wrapper{overflow-x:auto;}",
  "table.log-table{border-collapse:collapse;font-size:12px;}",
  "table.log-table th, table.log-table td{",
  "  padding:2px 4px;border:none;text-align:left;white-space:nowrap;font-family:monospace;}",
  "</style>",
  "</head><body>",
  '<div class="log-container">',
  funAll_return_REsult$logHtml,
  "</div>",
  "</body></html>"
)

myHtml_content <- file(outPutFileNameHtml, "w", encoding = "UTF-8")
writeLines(full_html, myHtml_content)
close(myHtml_content)

print("----------------- 程序结束 ---------------------")

