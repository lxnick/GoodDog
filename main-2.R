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

abort_with_message <- function(title = "Abort", detail = NULL) {
  cat("\nAbort Execution:\n\t", detail)
  flush.console()
  stop(title)
}

# Step 1 Check
# cat("\nStep 1, Start execution...\n")
# abort_with_message("Pause Step 1")

# 2. 讀入資料（請修改為你的實際資料路徑）

home_data   <- read_excel("data/家養犬_填補後資料062301.xlsx")
shelter_data<- read_excel("data/收容犬_填補後資料062501.xlsx")

# 查看資料結構與行數
# str(home_data)
# str(shelter_data)
# nrow(home_data)
# nrow(shelter_data)
#abort_with_message("Pause Step 2-1")

# 查看哪些欄位找不到
behavior_vars <- c("train_and_obey","aggresive","fear_and_anxiety","separate","exciment","attachment",
                   "run","activity","behavior79","behavior80","behavior81","behavior82","behavior83",
                   "behavior84","behavior85","behavior86","behavior87","behavior88","behavior89",
                   "behavior90","behavior91","behavior92","behavior93","behavior94","behavior95",
                   "behavior96","behavior97","behavior98","behavior99","behavior100")
# setdiff(behavior_vars, colnames(home_data))
# setdiff(behavior_vars, colnames(shelter_data))
# character(0)
# character(0)
# abort_with_message("Pause Step 2-2")

# 3. 資料前處理
home_data[home_data == "N"]    <- NA
shelter_data[shelter_data == "N"] <- NA

summary(home_data)
sum(is.na(home_data))
summary(shelter_data)
sum(is.na(shelter_data))
# abort_with_message("Pause Step 3-1")

# 因子/數值轉換範例 註 餵食次數主要為1、2我認為影響非線性，故用類別result <- analyze_behavior("train_and_obey")
home_factor_vars <- c("home_alone_situation", "if_else_pet","people_group", "people_where", "people_sex", "dog_sex", "ligation", "species","feed_times","when_feed", "feed_what")
home_numeric_vars<- c("how_long_feed", "walk_dog_per_week","people_age", "dog_weight", "dog_age")
home_data[home_factor_vars]    <- lapply(home_data[home_factor_vars], as.factor)
home_data[home_numeric_vars]   <- lapply(home_data[home_numeric_vars], as.numeric)
#str(home_data)
#summary(home_data)
#abort_with_message("Pause Step 3-2")

shelter_factor_vars <- c("move_size","clean_frequent","people_where","people_group", "people_sex", "dog_sex", "ligation", "species","feed_times","when_feed", "feed_what")
shelter_numeric_vars<-c("how_long_here","social_with_human","social_with_dog","how_many_roommate","people_age", "dog_weight", "dog_age")
shelter_data[shelter_factor_vars] <- lapply(shelter_data[shelter_factor_vars], as.factor)
shelter_data[shelter_numeric_vars] <- lapply(shelter_data[shelter_numeric_vars], as.numeric)
#str(shelter_data)
#summary(shelter_data)
#abort_with_message("Pause Step 3-3")

# 4. 定義函式：對單一行為建立三套模型並回傳整理結果（含樣本數與隨機效應變異）
analyze_behavior <- function(behavior) {
  cat("\n--- 開始處理行為：", behavior, "---\n")
  
  co2 <- co3 <- NULL
  n2 <- n3 <- NA
  vc3 <- NULL
  vf2_df <- vf3_df <- NULL
  
  mod2_success <- FALSE
  mod3_success <- FALSE
#  abort_with_message("Pause Step 4-1")
  
  tryCatch({
    
    # 模型 2: 家養犬（lm）
    home_data[[behavior]] <- as.numeric(home_data[[behavior]])
#    abort_with_message("Pause Step 4-2-1")
    home_model_vars <- c(behavior, "home_alone_situation", "if_else_pet","people_group", "people_where", "people_sex", "dog_sex", "species","feed_times","when_feed", "feed_what","how_long_feed", "walk_dog_per_week","people_age", "dog_weight", "dog_age")
    dat2 <- na.omit(home_data[, intersect(colnames(home_data), home_model_vars)])
    cat("  模型 2 (home_data): na.omit 後的樣本數 =", nrow(dat2), "\n")
    
#    abort_with_message("Pause Step 4-2-2")
    
    if (nrow(dat2) >= 10) {
      fixed_effects_mod2_potential <- c("home_alone_situation", "if_else_pet","people_group", "people_sex", "dog_sex", "species", "feed_what", "walk_dog_per_week","people_age", "dog_weight")
      valid_fixed_effects_mod2 <- Filter(function(fe) {
        fe %in% colnames(dat2) &&
          ((is.numeric(dat2[[fe]]) && sd(dat2[[fe]]) > 1e-6) ||
             (is.factor(dat2[[fe]]) && length(levels(dat2[[fe]])) > 1))
      }, fixed_effects_mod2_potential)
      
      if (length(valid_fixed_effects_mod2) > 0) {
        mod2 <- tryCatch({
          formula_str_mod2 <- paste0(behavior, " ~ ", paste(valid_fixed_effects_mod2, collapse = " + "))
          lm(as.formula(formula_str_mod2), data = dat2)
        }, error = function(e) {
          cat("    !!! 模型 2 建立失敗: ", e$message, "\n"); NULL
        })
        
        if (!is.null(mod2)) {
          mod2_success <- TRUE
          co2 <- tryCatch({
            broom::tidy(mod2) %>% mutate(model = "home_data (lm)")
          }, error = function(e) NULL)
          n2 <- nobs(mod2)
          
          if (length(valid_fixed_effects_mod2) > 1) {
            vf2_df <- tryCatch({
              vif_vals <- car::vif(mod2)
              data.frame(term = names(vif_vals), VIF = as.numeric(vif_vals), model = "home_data (lm)")
              
            }, error = function(e) NULL)
          }
        }
      }
    } else {
      cat("  模型 2 (home_data): 樣本不足 (<10)，跳過。\n")
    }
    
    # 模型 3: 收容犬（lmer）
    cat("  模型 3 with lmer");
    shelter_data[[behavior]] <- as.numeric(shelter_data[[behavior]])
    shelter_model_vars <- c(behavior, "move_size","clean_frequent","people_group", "people_where", "people_sex", "dog_sex", "species","feed_times","when_feed", "feed_what","how_long_here","social_with_human","social_with_dog","how_many_roommate","people_age", "dog_weight", "dog_age")
    # 檢查欄位名稱
    setdiff(shelter_model_vars, colnames(shelter_data))
    
    dat3 <- na.omit(shelter_data[, intersect(colnames(shelter_data), shelter_model_vars)])
    cat("  模型 3 (shelter_data): na.omit 後的樣本數 =", nrow(dat3), "\n")
    
    if (nrow(dat3) >= 10 && length(unique(dat3$people_where)) >= 2) {
      fixed_effects_mod3_potential <- c("move_size","clean_frequent","people_group","people_sex", "dog_sex", "species", "feed_what","social_with_human","social_with_dog","how_many_roommate","people_age", "dog_weight")
  
        valid_fixed_effects_mod3 <- Filter(function(fe) {
        fe %in% colnames(dat3) &&
          ((is.numeric(dat3[[fe]]) && sd(dat3[[fe]]) > 1e-6) ||
             (is.factor(dat3[[fe]]) && length(levels(dat3[[fe]])) > 1))
      }, fixed_effects_mod3_potential)
         
      
      if (length(valid_fixed_effects_mod3) > 0) {
        mod3 <- tryCatch({
          formula_str_mod3 <- paste0(behavior, " ~ ", paste(valid_fixed_effects_mod3, collapse = " + "), " + (1|people_where)")
          lme4::lmer(as.formula(formula_str_mod3), data = dat3)
        }, error = function(e) {
          cat("    !!! 模型 3 建立失敗: ", e$message, "\n"); NULL
        })
        
        if (!is.null(mod3)) {
          mod3_success <- TRUE
          co3 <- tryCatch({
            broom.mixed::tidy(mod3, effects = "fixed") %>% mutate(model = "shelter_data (lmer)")
          }, error = function(e) NULL)
          n3 <- nobs(mod3)
          vc3 <- tryCatch({
            as.data.frame(VarCorr(mod3)) %>%
              dplyr::select(grp, vcov) %>%
              mutate(ICC = ifelse(grp == "people_group", vcov / (vcov + attr(VarCorr(mod3), "sc")^2), NA),
                     model = "shelter_data (lmer)",
                     section = "Random Effects")
          }, error = function(e) NULL)
          
           if (length(valid_fixed_effects_mod3) > 1) {

            vf3_df <- tryCatch({
            vif_vals <- performance::check_collinearity(mod3)
            print(vif_vals)
            # 檢查欄位是否都存在
#            if (all(c("Parameter", "VIF") %in% colnames(vif_vals))) {
            if (all(c("Term", "VIF") %in% colnames(vif_vals))) {            
              cat("    Trace 06\n")
#              vf <- as.data.frame(vif_vals)[, c("Parameter", "VIF")]
              vf <- as.data.frame(vif_vals)[, c("Term", "VIF")]              
              colnames(vf) <- c("Term", "VIF")
              vf$model <- "shelter_data (lmer)"
              vf
            } else {
              stop("check_collinearity() 回傳缺少必要欄位")
            }
          }, error = function(e) {
            cat("    !!! 模型 3 VIF 失敗: ", e$message, "\n")
            NULL
          })            
          }
        }
      }
    } else {
      cat("  模型 3 (shelter_data): 樣本不足或群組數不足，跳過。\n")
    }
    
    # 整理輸出
    df_n <- data.frame(
      model = c("home_data (lm)", "shelter_data (lmer)"),
      used_n = c(if (mod2_success) n2 else NA, if (mod3_success) n3 else NA),
      section = "Used N"
    )
    
    df_coefs <- bind_rows(
      if (!is.null(co2) && nrow(co2) > 0) co2 else NULL,
      if (!is.null(co3) && nrow(co3) > 0) co3 else NULL
    ) %>% mutate(section = "Fixed Effects")
    
    df_vif <- bind_rows(
      if (!is.null(vf2_df) && nrow(vf2_df) > 0) vf2_df else NULL,
      if (!is.null(vf3_df) && nrow(vf3_df) > 0) vf3_df else NULL
    )
    if (is.null(df_vif) || nrow(df_vif) == 0) {
      df_vif <- data.frame(term = character(0), VIF = numeric(0), model = character(0), section = character(0))
    } else {
      df_vif <- df_vif %>% mutate(section = "VIF")
    }
    
    df_random <- bind_rows(
      if (!is.null(vc3) && nrow(vc3) > 0) vc3 else NULL
    )
    if (is.null(df_random) || nrow(df_random) == 0) {
      df_random <- data.frame(grp = character(0), vcov = numeric(0), ICC = numeric(0), model = character(0), section = character(0))
    }
    
    final_result <- bind_rows(df_n, df_coefs, df_vif, df_random)
    cat("--- 處理行為：", behavior, "完成，返回", nrow(final_result), "行結果 ---\n")
    return(final_result)
    
  }, error = function(e) {
    cat("！！！行為變項", behavior, "的 analyze_behavior 函數遇到未預期錯誤：", e$message, "！！！\n")
    return(NULL)
  })
}
#no execution ...
#abort_with_message("Pause Step 4")

# 5. 定義行為變項名稱
behavior_vars <- c("train_and_obey","aggresive","fear_and_anxiety","separate","exciment","attachment","run","activity","behavior79","behavior80","behavior81","behavior82","behavior83","behavior84","behavior85","behavior86","behavior87","behavior88","behavior89","behavior90","behavior91","behavior92","behavior93","behavior94","behavior95","behavior96","behavior97","behavior98","behavior99","behavior100")
#no execution ...
#abort_with_message("Pause Step 4")

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
saveWorkbook(wb, file = "data/0627_01_UsedN_with_VIF_integrated.xlsx", overwrite = TRUE)
cat("\n所有行為變項處理完成。請檢查 Excel 檔案和 R Console 輸出。\n")



