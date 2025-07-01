# RScript for:
#   
#  
# Console:
# install.packages("lme4")
# install.packages(c("lme4", "readxl", "openxlsx", "broom.mixed", "broom", "dplyr", "car", "lmerTest"))
# install.packages("performance")
library(readxl)       # 讀取 Excel
library(lme4)         # 建立混合模型
library(openxlsx)     # 輸出 Excel（可寫多個工作表）
library(broom.mixed)  # 整理 lmer 模型（tidy）
library(broom)        # 整理 lm 模型（tidy）
library(dplyr)        # 資料處理
library(car)          # 計算 VIF（適用於 lm 和 lmer 模型）
library(lmerTest)     # 計算 p value

# 自訂 car:::vif.mer 的版本，處理 lmer 模型的 VIF
vif.mer <- function(mod) {
  sm <- summary(mod, correlation = TRUE)
  corF <- sm$coefficients[, 2]
  if (is.null(sm$correlation)) stop("模型缺少相關矩陣，無法計算 VIF")
  corM <- sm$correlation
  v <- diag(solve(corM))
  names(v) <- rownames(sm$coefficients)
  return(v)
}

# 2. 讀入資料（請修改為你的實際資料路徑）

home_data   <- read_excel("data/家養犬_填補後資料062301.xlsx")
shelter_data<- read_excel("data/收容犬_填補後資料062501.xlsx")

# 查看資料結構與行數
str(home_data)
str(shelter_data)
nrow(home_data)
nrow(shelter_data)

# 查看哪些欄位找不到
behavior_vars <- c("train_and_obey","aggresive","fear_and_anxiety","separate","exciment","attachment",
                   "run","activity","behavior79","behavior80","behavior81","behavior82","behavior83",
                   "behavior84","behavior85","behavior86","behavior87","behavior88","behavior89",
                   "behavior90","behavior91","behavior92","behavior93","behavior94","behavior95",
                   "behavior96","behavior97","behavior98","behavior99","behavior100")
setdiff(behavior_vars, colnames(home_data))
setdiff(behavior_vars, colnames(shelter_data))
#  character(0)
#  character(0)

# 3. 資料前處理
home_data[home_data == "N"]    <- NA
shelter_data[shelter_data == "N"] <- NA

summary(home_data)
sum(is.na(home_data))
summary(shelter_data)
sum(is.na(shelter_data))

# 因子/數值轉換範例 註 餵食次數主要為1、2我認為影響非線性，故用類別result <- analyze_behavior("train_and_obey")
home_factor_vars <- c("home_alone_situation", "if_else_pet","people_group", "people_where", "people_sex", "dog_sex", "ligation", "species","feed_times","when_feed", "feed_what")
home_numeric_vars<- c("how_long_feed", "walk_dog_per_week","people_age", "dog_weight", "dog_age")
home_data[home_factor_vars]    <- lapply(home_data[home_factor_vars], as.factor)
home_data[home_numeric_vars]   <- lapply(home_data[home_numeric_vars], as.numeric)
str(home_data)
summary(home_data)

shelter_factor_vars <- c("move_size","clean_frequent","people_where","people_group", "people_sex", "dog_sex", "ligation", "species","feed_times","when_feed", "feed_what")
shelter_numeric_vars<-c("how_long_here","social_with_human","social_with_dog","how_many_roommate","people_age", "dog_weight", "dog_age")
shelter_data[shelter_factor_vars] <- lapply(shelter_data[shelter_factor_vars], as.factor)
shelter_data[shelter_numeric_vars] <- lapply(shelter_data[shelter_numeric_vars], as.numeric)
str(shelter_data)
summary(shelter_data)

# 測試執行函式
# analyze_behavior <- function(behavior) {
# cat("測試函式成功：你要分析的行為是：", behavior, "\n")
# }
#
# exists("analyze_behavior")
#


# 4. 定義函式：對單一行為建立三套模型並回傳整理結果（含樣本數與隨機效應變異）
analyze_behavior <- function(behavior) {
  cat("\n--- 開始處理行為：", behavior, "---\n")
  
  
  # 初始化所有結果變數，即使模型不運行，它們也是空的，避免後續bind_rows報錯
  co2 <- co3 <- NULL
  n2 <- n3 <- NA
  vc3 <- NULL
  vf2_df <- vf3_df <- NULL
  
  mod2_success <- FALSE
  mod3_success <- FALSE
  
  tryCatch({ # 外層 tryCatch 捕捉 analyze_behavior 函數中的一般性錯誤
    
    # 2) 家養犬 lm 模型
    home_data[[behavior]] <- as.numeric(home_data[[behavior]])
    home_model_vars <- c(behavior, "home_alone_situation", "if_else_pet","people_group", "people_where", "people_sex", "dog_sex", "species","feed_times","when_feed", "feed_what","how_long_feed", "walk_dog_per_week","people_age", "dog_weight", "dog_age")
    dat2 <- na.omit(home_data[, intersect(colnames(home_data), home_model_vars)])
    cat("  模型 2 (home_data): na.omit 後的樣本數 =", nrow(dat2), "\n")
    if (nrow(dat2) < 10) {
      cat("  模型 2 (home_data): 樣本不足 (<10)，跳過。\n")
    } else {
      fixed_effects_mod2_potential <- c( "home_alone_situation", "if_else_pet","people_group", "people_sex", "dog_sex", "species", "feed_what", "walk_dog_per_week","people_age", "dog_weight")
      valid_fixed_effects_mod2 <- c()
      for (fe in fixed_effects_mod2_potential) {
        if (fe %in% colnames(dat2)) {
          if (is.numeric(dat2[[fe]])) {
            if (sd(dat2[[fe]], na.rm = TRUE) > 1e-6) {
              valid_fixed_effects_mod2 <- c(valid_fixed_effects_mod2, fe)
            }
          } else if (is.factor(dat2[[fe]])) {
            if (length(levels(factor(dat2[[fe]]))) > 1) {
              valid_fixed_effects_mod2 <- c(valid_fixed_effects_mod2, fe)
            }
          }
        }
      }
      
      if (length(valid_fixed_effects_mod2) == 0) {
        cat("  模型 2 (home_data): 沒有足夠的有效固定效應預測變數來建立模型，跳過。\n")
      } else {
        mod2 <- tryCatch({
          formula_str_mod2 <- paste0(behavior, " ~ ", paste(valid_fixed_effects_mod2, collapse = " + "))
          lm(as.formula(formula_str_mod2), data = dat2)
        }, error = function(e) {
          cat("    !!! 模型 2 (home_data) 建立失敗: ", e$message, "\n")
          return(NULL)
        })
        if (!is.null(mod2)) {
          cat("  [診斷] 模型 2 建立成功，開始分析 model.matrix...\n")
          X_mat2 <- tryCatch({
            model.matrix(formula(mod2), data = dat2)
          }, error = function(e) {
            cat("    !!! model.matrix 建立失敗: ", e$message, "\n")
            return(NULL)
          })
          
          if (!is.null(X_mat2)) {
            na_cols <- which(colSums(is.na(X_mat2)) > 0)
            zero_cols <- which(apply(X_mat2, 2, function(col) all(col == 0)))
            if (length(na_cols) > 0) {
              cat("    ⚠️  設計矩陣中有 NA 欄位: ", paste(colnames(X_mat2)[na_cols], collapse = ", "), "\n")
            }
            if (length(zero_cols) > 0) {
              cat("    ⚠️  設計矩陣中有全為 0 的欄位: ", paste(colnames(X_mat2)[zero_cols], collapse = ", "), "\n")
            }
          }
        }
        
        if (!is.null(mod2)) {
          mod2_success <- TRUE
          co2 <- tryCatch({
            broom::tidy(mod2) %>% mutate(model = "home_data (lm)")
          }, error = function(e) {
            cat("    !!! 提取模型 2 (home_data) 固定效應結果失敗: ", e$message, "\n")
            return(NULL)
          })
          n2 <- nobs(mod2)
          if (length(valid_fixed_effects_mod2) > 1) {
            vf2_df <- tryCatch({
              cat("    [VIF DEBUG] 模型 2 行為：", behavior, "\n")
              cat("    [VIF DEBUG] 模型 2 類型：", class(mod2), "\n")
              X_mat <- model.matrix(mod2)
              cat("    [VIF DEBUG] model.matrix 維度：", paste(dim(X_mat), collapse = " x "), "\n")
              if (ncol(X_mat) <= 1) {
                stop("模型 2 的設計矩陣沒有足夠的變數，僅包含截距")
              }
              vf2 <- car::vif(mod2)
              data.frame(term = names(vf2), VIF = as.numeric(vf2), model = "home_data (lm)")
            }, error = function(e) {
              cat("    !!! 計算模型 2 (home_data) VIF 失敗: ", e$message, "\n")
              return(NULL)
            })
            
            
            
          } else {
            cat("  模型 2 (home_data): 固定效應預測變數少於2，不計算 VIF。\n")
          }
        }
      }
    }
    
    # 3) 收容犬 lmer 模型
    shelter_data[[behavior]] <- as.numeric(shelter_data[[behavior]])
    shelter_model_vars <- c(behavior, "move_size","clean_frequent","people_group", "people_where", "people_sex", "dog_sex", "species","feed_times","when_feed", "feed_what","how_long_here","social_with_human","social_with_dog","how_many_roommate","people_age", "dog_weight", "dog_age"
    )
    dat3 <- na.omit(shelter_data[, intersect(colnames(shelter_data), shelter_model_vars)])
    cat("  模型 3 (shelter_data): na.omit 後的樣本數 =", nrow(dat3), "\n")
    if (nrow(dat3) < 10) {
      cat("  模型 3 (shelter_data): 樣本不足 (<10)，跳過。\n")
    } else {
      if (length(unique(dat3$people_where)) < 2) {
        cat("  模型 3 (shelter_data): people_where 隨機效應組別少於2，模型可能無法建立，跳過。\n")
      } else {
        fixed_effects_mod3_potential <- c("move_size","clean_frequent","people_group","people_sex", "dog_sex", "species", "feed_what","social_with_human","social_with_dog","how_many_roommate","people_age", "dog_weight"
        )
        valid_fixed_effects_mod3 <- c()
        
        for (fe in fixed_effects_mod3_potential) {
          if (fe %in% colnames(dat3)) {
            if (is.numeric(dat3[[fe]])) {
              if (sd(dat3[[fe]], na.rm = TRUE) > 1e-6) {
                valid_fixed_effects_mod3 <- c(valid_fixed_effects_mod3, fe)
              }
            } else if (is.factor(dat3[[fe]])) {
              if (length(levels(factor(dat3[[fe]]))) > 1) {
                valid_fixed_effects_mod3 <- c(valid_fixed_effects_mod3, fe)
              }
            }
          }
        }
        
        if (length(valid_fixed_effects_mod3) == 0) {
          cat("  模型 3 (shelter_data): 沒有足夠的有效固定效應預測變數來建立模型，跳過。\n")
        } else {
#          mod3 <- tryCatch({
#            formula_str_mod3 <- paste0(behavior, " ~ ", paste(valid_fixed_effects_mod3, collapse = " + "), " + (1|people_where)")
#            lmer(as.formula(formula_str_mod3), data = dat3)
#          }, error = function(e) {
#            cat("    !!! 模型 3 (shelter_data) 建立失敗: ", e$message, "\n")
#            return(NULL)
#          })
          mod3 <- tryCatch({
            formula_str_mod3 <- paste0(behavior, " ~ ", paste(valid_fixed_effects_mod3, collapse = " + "), " + (1|people_where)")
            lme4::lmer(as.formula(formula_str_mod3), data = dat3)
          }, error = function(e) {
            cat("    !!! 模型 3 (shelter_data) 建立失敗: ", e$message, "\n")
            return(NULL)
          })          
          
          if (!is.null(mod3)) {
            mod3_success <- TRUE
            co3 <- tryCatch({
              broom.mixed::tidy(mod3, effects = "fixed") %>% mutate(model = "shelter_data (lmer)")
            }, error = function(e) {
              cat("    !!! 提取模型 3 (shelter_data) 固定效應結果失敗: ", e$message, "\n")
              return(NULL)
            })
            n3 <- nobs(mod3)
            vc3 <- tryCatch({
              as.data.frame(VarCorr(mod3)) %>%
                dplyr::select(grp, vcov) %>%
                mutate(ICC = ifelse(grp == "people_group", vcov / (vcov + attr(VarCorr(mod3), "sc")^2), NA),
                       model = "shelter_data (lmer)",
                       section = "Random Effects")
            }, error = function(e) {
              cat("    !!! 提取模型 3 (shelter_data) 隨機效應變異失敗: ", e$message, "\n")
              return(NULL)
            })
            
            if (length(valid_fixed_effects_mod3) > 1) {
              vf3_df <- tryCatch({
                cat("    [VIF DEBUG] 模型 3 類型：", class(mod3), "\n")
                
                # 確保 summary 有 correlation 矩陣
                s_mod3 <- summary(mod3, correlation = TRUE)
                
                if (is.null(s_mod3$correlation)) {
                  stop("summary(mod3)$correlation 是 NULL，無法計算 VIF")
                }
                
                # car::vif 會用 summary(model)$correlation 來計算
                # vf3 <- car:::vif.mer(mod3)
                # vf3 <- vif.mer(mod3)
                vf3 <- performance::check_collinearity(mod3)
                vf3 <- as.data.frame(vf3)
                vf3 <- vf3[, c("Parameter", "VIF")]  # 篩出主要欄位
                colnames(vf3) <- c("term", "VIF")
                vf3$model <- "shelter_data (lmer)"
                
                data.frame(term = names(vf3), VIF = as.numeric(vf3), model = "shelter_data (lmer)")
              }, error = function(e) {
                cat("    !!! 計算模型 3 (shelter_data) VIF 失敗: ", e$message, "\n")
                return(NULL)
              })
              
            } else {
              cat("  模型 3 (shelter_data): 固定效應預測變數少於2，不計算 VIF。\n")
            }
          }
        }
      }
    }
    
    # 實際用到樣本數表格 (只有在模型成功建立的情況下才使用其nobs)
    df_n <- data.frame(
      model  = c("home_data (lm)", "shelter_data (lmer)"),
      used_n = c(if(mod2_success) n2 else NA, if(mod3_success) n3 else NA),
      section = "Used N"
    )
    
    # 固定效應 (使用 if (!is.null(coX) && nrow(coX) > 0) 確保只合併有資料的結果)
    df_coefs <- bind_rows(
      if (!is.null(co2) && nrow(co2) > 0) co2 else NULL,
      if (!is.null(co3) && nrow(co3) > 0) co3 else NULL
    ) %>% mutate(section = "Fixed Effects")
    
    # VIF (使用 if (!is.null(vfX_df) && nrow(vfX_df) > 0) 確保只合併有資料的結果)
    df_vif   <- bind_rows(
      if (!is.null(vf2_df) && nrow(vf2_df) > 0) vf2_df else NULL,
      if (!is.null(vf3_df) && nrow(vf3_df) > 0) vf3_df else NULL
    )
    if (is.null(df_vif) || nrow(df_vif) == 0) {
      df_vif <- data.frame(term = character(0), VIF = numeric(0), model = character(0), section = character(0))
    } else {
      df_vif <- df_vif %>% mutate(section = "VIF")
    }
    
    # 隨機效應 (使用 if (!is.null(vcX) && nrow(vcX) > 0) 確保只合併有資料的結果)
    df_random <- bind_rows(
      if (!is.null(vc3) && nrow(vc3) > 0) vc3 else NULL
    )
    if (is.null(df_random) || nrow(df_random) == 0) {
      df_random <- data.frame(grp = character(0), vcov = numeric(0), ICC = numeric(0), model = character(0), section = character(0))
    }
    
    
    # 回傳合併結果
    final_result <- bind_rows(df_n, df_coefs, df_vif, df_random)
    cat("--- 處理行為：", behavior, "完成，返回", nrow(final_result), "行結果 ---\n")
    return(final_result)
  }, error = function(e) { # 外層 tryCatch 捕捉 analyze_behavior 函數中的一般性錯誤
    cat("！！！行為變項", behavior, "的 analyze_behavior 函數遇到未預期錯誤：", e$message, "！！！\n")
    return(NULL)
  })
}


# 5. 定義行為變項名稱
behavior_vars <- c("train_and_obey","aggresive","fear_and_anxiety","separate","exciment","attachment","run","activity","behavior79","behavior80","behavior81","behavior82","behavior83","behavior84","behavior85","behavior86","behavior87","behavior88","behavior89","behavior90","behavior91","behavior92","behavior93","behavior94","behavior95","behavior96","behavior97","behavior98","behavior99","behavior100")

# 6. 批次寫入 Excel
wb <- createWorkbook()
for (b in behavior_vars) {
  cat("正在處理行為：", b, "\n")
  res <- analyze_behavior(b)
  if (!is.null(res) && nrow(res)>0) {
    addWorksheet(wb, sheetName = b)
    writeData(wb, sheet = b, x = res)
  } else {
    cat("行為變項：", b, " - 未能產生有效的分析結果。\n")
  }
}
# 7. 儲存檔案
saveWorkbook(wb, file = "C:/Users/User/OneDrive/Desktop/作業資料(高中)/個研資料/0627_01_UsedN_with_VIF_integrated.xlsx", overwrite = TRUE)
cat("\n所有行為變項處理完成。請檢查 Excel 檔案和 R Console 輸出。\n")



