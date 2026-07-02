# =============================================================================
# 以 follow_data 为基础，抓取 function_raw_20260507_1 的 function 数据
#
# 规则：
#   - 以 follow_data（1213 例）为基准，保留全部行（左连接）
#   - 唯一匹配条件：patient_id（前导 0 视为同一 ID，0054288752 == 54288752）
#   - 从 function 表提取每行全部数据；follow_data 中已有的同名列不重复提取
#   - 某患者在 function 表中找不到 -> function 各列用 NA 占位，并整段标【绿色】
#   - 匹配上但两表 procedure_date 不一致 -> 该行 procedure_date 单元格标【红色】
#
# 输出：output/follow_with_function.xlsx（带颜色，用 openxlsx 写出）
# =============================================================================

# ---- 1. 依赖包（首次运行请先取消下一行注释安装）-----------------------------
# install.packages(c("readxl", "openxlsx", "dplyr"))
library(readxl)
library(openxlsx)
library(dplyr)

# ---- 2. 路径与参数（如列名或文件名不同，只改这里）---------------------------
data_dir   <- "D:/R_project/随访+OCT+功能/data"
output_dir <- "D:/R_project/随访+OCT+功能/output"

follow_file <- "follow_data.xlsx"
func_file   <- "function_raw_20260507_1.xlsx"
out_file    <- "follow_with_function.xlsx"

id_col   <- "patient_id"       # 匹配用 ID 列名（两张表须一致）
date_col <- "procedure_date"   # 日期列名（用于红色标注比对）

green_fill <- "#C6EFCE"        # 未匹配（NA 占位）标注色
red_fill   <- "#FFC7CE"        # 日期不一致标注色

# ---- 3. 读取两张表 -----------------------------------------------------------
follow <- read_excel(file.path(data_dir, follow_file))
func   <- read_excel(file.path(data_dir, func_file))

# ---- 4. 归一化函数 -----------------------------------------------------------
# patient_id：去掉前导 0，使 0054288752 与 54288752 视为同一 ID
norm_id <- function(x) {
  na <- is.na(x)
  if (is.numeric(x)) x <- sprintf("%.0f", x) else x <- as.character(x)
  x <- trimws(x)
  x <- sub("^0+", "", x)
  x[x == ""] <- "0"
  x[na] <- NA
  x
}
# procedure_date：统一成 "YYYY-MM-DD"，兼容日期型 / Excel序列号 / 文本
norm_date <- function(x) {
  if (inherits(x, "Date"))   return(format(x, "%Y-%m-%d"))
  if (inherits(x, "POSIXt")) return(format(as.Date(x), "%Y-%m-%d"))
  if (is.numeric(x))         return(format(as.Date(x, origin = "1899-12-30"), "%Y-%m-%d"))
  x <- trimws(as.character(x))
  d <- as.Date(x, tryFormats = c("%Y-%m-%d","%Y/%m/%d","%m/%d/%Y","%d/%m/%Y","%Y.%m.%d","%Y年%m月%d日"))
  format(d, "%Y-%m-%d")
}

follow$.pid <- norm_id(follow[[id_col]])
func$.pid   <- norm_id(func[[id_col]])

# ---- 5. function 表按 patient_id 去重（保留每个 ID 第一条），并挑选要追加的列 --
dups <- sum(duplicated(func$.pid))
if (dups > 0)
  message(sprintf("注意：function 表存在 %d 个重复 patient_id，已保留每个 ID 的第一条记录用于匹配。", dups))
func <- func[!duplicated(func$.pid), ]

# 要追加的列 = function 列中、follow_data 里尚不存在的同名列（且排除 key）
add_cols <- setdiff(names(func), c(names(follow), ".pid"))

# function 表参与匹配的子集：key + 追加列 + 匹配标记 + 用于比对的 function 日期
func_sel <- func[, c(".pid", add_cols), drop = FALSE]
func_sel$.matched <- TRUE
has_func_date <- date_col %in% names(func)
if (has_func_date) func_sel$.func_date <- func[[date_col]]

# ---- 6. 左连接（follow 为基准，保留全部行）----------------------------------
result <- follow %>% left_join(func_sel, by = ".pid")
result$.matched[is.na(result$.matched)] <- FALSE
matched_vec <- result$.matched

# 日期不一致（仅对匹配上、且两边日期都存在的行判断）
if (has_func_date && (date_col %in% names(result))) {
  fdate <- norm_date(result[[date_col]])
  gdate <- norm_date(result$.func_date)
  mismatch_vec <- matched_vec & !is.na(fdate) & !is.na(gdate) & (fdate != gdate)
} else {
  mismatch_vec <- rep(FALSE, nrow(result))
}

# ---- 7. 整理输出表（去掉辅助列，不重复输出 function 的日期列）----------------
out <- result %>% select(-any_of(c(".pid", ".func_date", ".matched")))

func_col_idx <- match(add_cols, names(out))          # 追加的 function 列位置
func_col_idx <- func_col_idx[!is.na(func_col_idx)]
date_col_idx <- match(date_col, names(out))          # procedure_date 列位置

# ---- 8. 汇总信息 -------------------------------------------------------------
cat("follow_data 行数：      ", nrow(follow), "\n")
cat("成功匹配 function：     ", sum(matched_vec), "\n")
cat("未匹配（绿色标注）：    ", sum(!matched_vec), "\n")
cat("日期不一致（红色标注）：", sum(mismatch_vec), "\n")
cat("输出列数：              ", ncol(out), "\n")

# ---- 9. 用 openxlsx 写出并着色 ----------------------------------------------
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

wb <- createWorkbook()
addWorksheet(wb, "follow_function")
# keepNA=TRUE 让未匹配的 function 单元格显示 "NA" 占位
writeData(wb, "follow_function", out, keepNA = TRUE, na.string = "NA")

greenStyle <- createStyle(fgFill = green_fill)
redStyle   <- createStyle(fgFill = red_fill)

# 绿色：未匹配行的所有 function 追加列（NA 占位）
unmatched_rows <- which(!matched_vec)
if (length(unmatched_rows) > 0 && length(func_col_idx) > 0) {
  addStyle(wb, "follow_function", greenStyle,
           rows = unmatched_rows + 1, cols = func_col_idx,   # +1 跳过表头
           gridExpand = TRUE, stack = TRUE)
}

# 红色：日期不一致行的 procedure_date 单元格
mismatch_rows <- which(mismatch_vec)
if (length(mismatch_rows) > 0 && !is.na(date_col_idx)) {
  addStyle(wb, "follow_function", redStyle,
           rows = mismatch_rows + 1, cols = date_col_idx,
           gridExpand = TRUE, stack = TRUE)
}

out_path <- file.path(output_dir, out_file)
saveWorkbook(wb, out_path, overwrite = TRUE)
cat("已生成：", out_path, "\n")
