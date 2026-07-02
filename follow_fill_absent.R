# =============================================================================
# 以 follow_function_oct 为基础，从 function_raw_absent_20260607 补齐缺失的功能学数据
#
# 背景：follow_function_oct 中 id / Target_vessel / angio_ffr / angio_imr /
#       measurement_angle 这 5 列已存在，但部分患者为 NA（缺失）。
#
# 规则：
#   - 目标行 = base 中 angio_ffr 且 angio_imr 均为 NA 的行（要补齐的行）
#   - 数据源 = function_raw_absent 中【仅 angio_ffr、angio_imr 均为 NA】的行
#     （该表里也有 ffr/imr 有值的行，必须排除）
#   - 唯一匹配条件：patient_id（前导 0 视为同一 ID，0054288752 == 54288752）
#   - 只【填补 NA 单元格】，base 已有值的不覆盖（不重复提取）
#   - 目标行仍未在数据源中找到 -> 保持 NA，并把仍为 NA 的 5 列单元格标【浅蓝色】
#   - 匹配上但两表 procedure_date 不一致 -> 该行 procedure_date 单元格标【红色】
#
# 注意：前几步把空值写成了文字 "NA"，读入时用 na=c("","NA") 还原为真正的 NA。
#
# 输出：output/follow_function_oct_filled.xlsx（带颜色，用 openxlsx 写出）
# =============================================================================

# ---- 1. 依赖包（首次运行请先取消下一行注释安装）-----------------------------
# install.packages(c("readxl", "openxlsx", "dplyr"))
library(readxl)
library(openxlsx)
library(dplyr)

# ---- 2. 路径与参数（如列名或文件名不同，只改这里）---------------------------
data_dir   <- "D:/R_project/随访+OCT+功能/data"
output_dir <- "D:/R_project/随访+OCT+功能/output"

base_file   <- "follow_function_oct.xlsx"              # 基准表（已放在 data 文件夹）
absent_file <- "function_raw_absent_20260607.xlsx"     # 补充数据源
out_file    <- "follow_function_oct_filled.xlsx"

id_col   <- "patient_id"       # 匹配用 ID 列名（两张表须一致）
date_col <- "procedure_date"   # 日期列名（用于红色标注比对）

# 要补齐的 5 列（base 与 absent 中同名）
fill_cols <- c("id", "Target_vessel", "angio_ffr", "angio_imr", "measurement_angle")
# 判定“缺失/目标行”依据的列：这些都为 NA 才算需要补齐
key_missing <- c("angio_ffr", "angio_imr")

blue_fill <- "#ADD8E6"         # 未匹配（仍为 NA）标注色（浅蓝）
red_fill  <- "#FFC7CE"         # 日期不一致标注色

# ---- 3. 读取（把 "" 和 "NA" 都当作真正的 NA）--------------------------------
base   <- read_excel(file.path(data_dir, base_file),   na = c("", "NA"))
absent <- read_excel(file.path(data_dir, absent_file), na = c("", "NA"))

# ---- 4. 归一化函数 -----------------------------------------------------------
norm_id <- function(x) {
  na <- is.na(x)
  if (is.numeric(x)) x <- sprintf("%.0f", x) else x <- as.character(x)
  x <- trimws(x); x <- sub("^0+", "", x); x[x == ""] <- "0"; x[na] <- NA
  x
}
norm_date <- function(x) {
  if (inherits(x, "Date"))   return(format(x, "%Y-%m-%d"))
  if (inherits(x, "POSIXt")) return(format(as.Date(x), "%Y-%m-%d"))
  if (is.numeric(x))         return(format(as.Date(x, origin = "1899-12-30"), "%Y-%m-%d"))
  x <- trimws(as.character(x))
  format(as.Date(x, tryFormats = c("%Y-%m-%d","%Y/%m/%d","%m/%d/%Y","%d/%m/%Y","%Y.%m.%d","%Y年%m月%d日")),
         "%Y-%m-%d")
}

base$.pid   <- norm_id(base[[id_col]])
absent$.pid <- norm_id(absent[[id_col]])

# ---- 5. 检查列 & 统一为字符型（便于 NA 判定与回填）--------------------------
base_fill    <- intersect(fill_cols, names(base))                       # base 中用于标注的列
present_fill <- intersect(fill_cols, intersect(names(base), names(absent)))  # 两表都有 -> 可回填
if (length(setdiff(fill_cols, names(base)))   > 0)
  message("提示：base 缺少列：",   paste(setdiff(fill_cols, names(base)),   collapse=", "))
if (length(setdiff(fill_cols, names(absent))) > 0)
  message("提示：absent 缺少列：", paste(setdiff(fill_cols, names(absent)), collapse=", "))
stopifnot(all(key_missing %in% names(base)), all(key_missing %in% names(absent)))

for (col in base_fill)    base[[col]]   <- as.character(base[[col]])
for (col in present_fill) absent[[col]] <- as.character(absent[[col]])

# ---- 6. 目标行 & 数据源筛选 --------------------------------------------------
# 目标行：base 中 key_missing 全部为 NA
target <- Reduce(`&`, lapply(key_missing, function(c) is.na(base[[c]])))

# 数据源：absent 中 key_missing 全部为 NA 的行
absent_na <- Reduce(`&`, lapply(key_missing, function(c) is.na(absent[[c]])))
absent_f  <- absent[absent_na, , drop = FALSE]

dups <- sum(duplicated(absent_f$.pid))
if (dups > 0)
  message(sprintf("注意：absent 表(仅 ffr/imr 均为 NA 的行)中有 %d 个重复 patient_id，已保留每个 ID 首条。", dups))
absent_f <- absent_f[!duplicated(absent_f$.pid), ]

# 查找表：.pid + 回填列（改名避免冲突）+ 匹配标记 + 用于比对的日期
lookup <- absent_f[, ".pid", drop = FALSE]
for (col in present_fill) lookup[[paste0(col, "__fill")]] <- absent_f[[col]]
lookup$.matched <- TRUE
has_absent_date <- date_col %in% names(absent_f)
if (has_absent_date) lookup$.absent_date <- absent_f[[date_col]]

# ---- 7. 左连接 + 回填（仅 目标行 & 匹配上 & base 该格为 NA）------------------
merged  <- base %>% left_join(lookup, by = ".pid")
matched <- !is.na(merged$.matched)

for (col in present_fill) {
  fillv   <- merged[[paste0(col, "__fill")]]
  do_fill <- target & matched & is.na(base[[col]])
  do_fill[is.na(do_fill)] <- FALSE
  base[[col]][do_fill] <- fillv[do_fill]
}

# 日期不一致（仅 目标行 & 匹配上 & 两边日期都存在）
if (has_absent_date && (date_col %in% names(merged))) {
  bdate <- norm_date(merged[[date_col]])
  adate <- norm_date(merged$.absent_date)
  mismatch <- target & matched & !is.na(bdate) & !is.na(adate) & (bdate != adate)
} else {
  mismatch <- rep(FALSE, nrow(base))
}

# ---- 8. 整理输出（不新增列，去掉辅助键）-------------------------------------
out <- base %>% select(-any_of(".pid"))
date_idx <- match(date_col, names(out))

# ---- 9. 汇总信息 -------------------------------------------------------------
cat("基准表 行数：              ", nrow(out), "\n")
cat("需补齐的目标行(ffr&imr NA)：", sum(target), "\n")
cat("其中匹配到数据源：         ", sum(target & matched), "\n")
cat("仍未匹配(浅蓝标注)：       ", sum(target & !matched), "\n")
cat("日期不一致(红色标注)：     ", sum(mismatch), "\n")

# ---- 10. 用 openxlsx 写出并着色 ---------------------------------------------
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

wb <- createWorkbook()
addWorksheet(wb, "filled")
writeData(wb, "filled", out, keepNA = TRUE, na.string = "NA")   # NA 显示为 "NA" 占位

blueStyle <- createStyle(fgFill = blue_fill)
redStyle  <- createStyle(fgFill = red_fill)

# 浅蓝：目标行中未匹配、且仍为 NA 的 5 列单元格（逐列应用）
for (col in base_fill) {
  ci <- match(col, names(out))
  if (is.na(ci)) next
  rows_blue <- which(target & !matched & is.na(out[[col]]))
  if (length(rows_blue) > 0)
    addStyle(wb, "filled", blueStyle, rows = rows_blue + 1, cols = ci,   # +1 跳过表头
             gridExpand = TRUE, stack = TRUE)
}

# 红色：日期不一致行的 procedure_date 单元格
mm_rows <- which(mismatch)
if (length(mm_rows) > 0 && !is.na(date_idx))
  addStyle(wb, "filled", redStyle, rows = mm_rows + 1, cols = date_idx,
           gridExpand = TRUE, stack = TRUE)

out_path <- file.path(output_dir, out_file)
saveWorkbook(wb, out_path, overwrite = TRUE)
cat("已生成：", out_path, "\n")
