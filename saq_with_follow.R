# =============================================================================
# 以 SAQ_scores_raw 为基础（235 名患者），从 follow_function_oct_filled 追加数据
#
# 规则：
#   - 以 SAQ_scores_raw 为基准，保留全部行（左连接）
#   - 唯一匹配条件：patient_name（姓名，去除空格后精确匹配）
#   - 匹配到就把源表该行的全部列追加过来；SAQ 已有的同名列不重复追加
#   - 某姓名在源表中找不到 -> 追加列用 NA 占位，并整段标【棕色】
#
# 注意：姓名可能重名。若源表中同名多行，默认保留第一条并打印警告——
#       可能取错人，请结合警告核对（如两表都有 patient_id，建议改用 id 匹配）。
#
# 输出：SAQ_output/SAQ_with_follow.xlsx
# =============================================================================

# ---- 1. 依赖包（首次运行请先取消下一行注释安装）-----------------------------
# install.packages(c("readxl", "openxlsx", "dplyr"))
library(readxl)
library(openxlsx)
library(dplyr)

# ---- 2. 路径与参数（如列名或文件名不同，只改这里）---------------------------
data_dir   <- "D:/R_project/OCT研究/SAQ_data"
output_dir <- "D:/R_project/OCT研究/SAQ_output"

base_file   <- "SAQ_scores_raw.xlsx"                # 基准表（235 名患者）
source_file <- "follow_function_oct_filled.xlsx"    # 数据源
out_file    <- "SAQ_with_follow.xlsx"

name_col <- "patient_name"     # 匹配用姓名列名（两张表须一致）

brown_fill <- "#B5651D"        # 未匹配（NA 占位）标注色（棕色，可改）

# ---- 3. 读取（"" 和 "NA" 都当作真正的 NA）--------------------------------
base <- read_excel(file.path(data_dir, base_file),   na = c("", "NA"))
src  <- read_excel(file.path(data_dir, source_file), na = c("", "NA"))

# ---- 4. 姓名归一化（去除所有空格，含全角空格；不改大小写）------------------
norm_name <- function(x) {
  x <- as.character(x)
  x <- gsub("[[:space:]]+", "", x)          # 去 ASCII 空白
  x <- gsub("　", "", x, fixed = TRUE)  # 去全角空格
  x
}
base$.name <- norm_name(base[[name_col]])
src$.name  <- norm_name(src[[name_col]])

# ---- 5. 源表按姓名去重（去掉空姓名；每个姓名保留第一条）----------------------
src <- src[!is.na(src$.name) & src$.name != "", , drop = FALSE]
dups <- sum(duplicated(src$.name))
if (dups > 0)
  message(sprintf("注意：源表中有 %d 个重复 patient_name，已保留每个姓名第一条（姓名匹配可能不唯一，请核对）。", dups))
src <- src[!duplicated(src$.name), ]

# 要追加的列 = 源表列中、SAQ 里尚不存在的同名列（且排除 key）
add_cols <- setdiff(names(src), c(names(base), ".name"))
src_j <- src[, c(".name", add_cols), drop = FALSE]
src_j$.matched <- TRUE

# ---- 6. 左连接（SAQ 为基准，保留全部行）------------------------------------
merged      <- base %>% left_join(src_j, by = ".name")
matched_vec <- !is.na(merged$.matched)

# ---- 7. 整理输出（去掉辅助列）-----------------------------------------------
out <- merged %>% select(-any_of(c(".name", ".matched")))
add_col_idx <- match(add_cols, names(out))
add_col_idx <- add_col_idx[!is.na(add_col_idx)]

# ---- 8. 汇总信息 -------------------------------------------------------------
cat("SAQ 行数：          ", nrow(base), "\n")
cat("成功匹配：          ", sum(matched_vec), "\n")
cat("未匹配（棕色标注）：", sum(!matched_vec), "\n")
cat("输出列数：          ", ncol(out), "\n")

# ---- 9. 用 openxlsx 写出并着色 ----------------------------------------------
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

wb <- createWorkbook()
addWorksheet(wb, "SAQ_follow")
writeData(wb, "SAQ_follow", out, keepNA = TRUE, na.string = "NA")   # NA 显示为 "NA" 占位

brownStyle <- createStyle(fgFill = brown_fill)

# 棕色：未匹配行的所有追加列（NA 占位）
unmatched_rows <- which(!matched_vec)
if (length(unmatched_rows) > 0 && length(add_col_idx) > 0) {
  addStyle(wb, "SAQ_follow", brownStyle,
           rows = unmatched_rows + 1, cols = add_col_idx,   # +1 跳过表头
           gridExpand = TRUE, stack = TRUE)
}

out_path <- file.path(output_dir, out_file)
saveWorkbook(wb, out_path, overwrite = TRUE)
cat("已生成：", out_path, "\n")
