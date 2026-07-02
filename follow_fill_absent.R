# =============================================================================
# 以 follow_function_oct 为基础，从 function_raw_absent_20260607 补齐缺失功能学数据
#
# 背景：follow_function_oct 中 id / Target_vessel / angio_ffr / angio_imr /
#       measurement_angle 这 5 列已存在，但部分患者为 NA（缺失）。
#       function_raw_absent 表里存有这些患者的实测值（重点是 ffr/imr/angle）。
#
# 规则：
#   - 目标行 = base 中 angio_ffr 且 angio_imr 均为 NA 的行（要补齐的行）
#   - 唯一匹配条件：patient_id（前导 0 视为同一 ID，0052837846 == 52837846）
#   - 匹配到就把 5 列的值填入 base 的 NA 单元格；base 已有值的不覆盖
#   - 若某 patient_id 在 absent 表里有多行，优先取 ffr/imr 有值的那一行
#   - 目标行仍未找到 -> 保持 NA，并把仍为 NA 的 5 列单元格标【浅蓝色】
#
# 注意：前几步把空值写成了文字 "NA"，读入时用 na=c("","NA") 还原为真正的 NA。
#
# 输出：output/follow_function_oct_filled.xlsx
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

id_col <- "patient_id"          # 匹配用 ID 列名（两张表须一致）

# 要补齐的 5 列（base 与 absent 中同名）
fill_cols <- c("id", "Target_vessel", "angio_ffr", "angio_imr", "measurement_angle")
# 判定“缺失/目标行”依据的列：这些都为 NA 才算需要补齐
key_missing <- c("angio_ffr", "angio_imr")

blue_fill <- "#ADD8E6"          # 未匹配（仍为 NA）标注色（浅蓝）

# ---- 3. 读取（把 "" 和 "NA" 都当作真正的 NA）--------------------------------
base   <- read_excel(file.path(data_dir, base_file),   na = c("", "NA"))
absent <- read_excel(file.path(data_dir, absent_file), na = c("", "NA"))

# ---- 4. 归一化 patient_id（去前导 0）----------------------------------------
norm_id <- function(x) {
  na <- is.na(x)
  if (is.numeric(x)) x <- sprintf("%.0f", x) else x <- as.character(x)
  x <- trimws(x); x <- sub("^0+", "", x); x[x == ""] <- "0"; x[na] <- NA
  x
}
base$.pid   <- norm_id(base[[id_col]])
absent$.pid <- norm_id(absent[[id_col]])

# ---- 5. 检查列 & 统一为字符型（便于 NA 判定与回填）--------------------------
base_fill    <- intersect(fill_cols, names(base))                            # base 中用于标注的列
present_fill <- intersect(fill_cols, intersect(names(base), names(absent)))  # 两表都有 -> 可回填
if (length(setdiff(fill_cols, names(base)))   > 0)
  message("提示：base 缺少列：",   paste(setdiff(fill_cols, names(base)),   collapse = ", "))
if (length(setdiff(fill_cols, names(absent))) > 0)
  message("提示：absent 缺少列：", paste(setdiff(fill_cols, names(absent)), collapse = ", "))
stopifnot(all(key_missing %in% names(base)), all(key_missing %in% names(absent)))

for (col in base_fill)    base[[col]]   <- as.character(base[[col]])
for (col in present_fill) absent[[col]] <- as.character(absent[[col]])

# ---- 6. 目标行 -------------------------------------------------------------
target <- Reduce(`&`, lapply(key_missing, function(c) is.na(base[[c]])))

# ---- 7. absent 去重：每个 patient_id 只留一行，优先留 ffr/imr 有值的那行 -----
have_val <- !is.na(absent[["angio_ffr"]]) & !is.na(absent[["angio_imr"]])
ord      <- order(-as.integer(have_val), seq_len(nrow(absent)))   # 有值的排前面（稳定）
absent_o <- absent[ord, , drop = FALSE]
dups <- sum(duplicated(absent_o$.pid))
if (dups > 0)
  message(sprintf("注意：absent 表中有 %d 个重复 patient_id，已保留每个 ID（优先有值）的一行。", dups))
absent_u <- absent_o[!duplicated(absent_o$.pid), ]

# 查找表：.pid + 回填列（改名避免冲突）+ 匹配标记
lookup <- absent_u[, ".pid", drop = FALSE]
for (col in present_fill) lookup[[paste0(col, "__fill")]] <- absent_u[[col]]
lookup$.matched <- TRUE

# ---- 8. 左连接 + 回填（仅 目标行 & 匹配上 & base 该格为 NA）------------------
merged  <- base %>% left_join(lookup, by = ".pid")
matched <- !is.na(merged$.matched)

for (col in present_fill) {
  fillv   <- merged[[paste0(col, "__fill")]]
  do_fill <- target & matched & is.na(base[[col]])
  do_fill[is.na(do_fill)] <- FALSE
  base[[col]][do_fill] <- fillv[do_fill]
}

# ---- 9. 整理输出（不新增列，去掉辅助键）-------------------------------------
out <- base %>% select(-any_of(".pid"))

# ---- 10. 汇总信息 ------------------------------------------------------------
cat("基准表 行数：               ", nrow(out), "\n")
cat("需补齐的目标行(ffr&imr NA)： ", sum(target), "\n")
cat("其中匹配到并已回填：        ", sum(target & matched), "\n")
cat("仍未匹配(浅蓝标注)：        ", sum(target & !matched), "\n")

# ---- 11. 用 openxlsx 写出并着色 ---------------------------------------------
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

wb <- createWorkbook()
addWorksheet(wb, "filled")
writeData(wb, "filled", out, keepNA = TRUE, na.string = "NA")   # NA 显示为 "NA" 占位

blueStyle <- createStyle(fgFill = blue_fill)

# 浅蓝：目标行中未匹配、且仍为 NA 的 5 列单元格（逐列应用）
for (col in base_fill) {
  ci <- match(col, names(out))
  if (is.na(ci)) next
  rows_blue <- which(target & !matched & is.na(out[[col]]))
  if (length(rows_blue) > 0)
    addStyle(wb, "filled", blueStyle, rows = rows_blue + 1, cols = ci,   # +1 跳过表头
             gridExpand = TRUE, stack = TRUE)
}

out_path <- file.path(output_dir, out_file)
saveWorkbook(wb, out_path, overwrite = TRUE)
cat("已生成：", out_path, "\n")
