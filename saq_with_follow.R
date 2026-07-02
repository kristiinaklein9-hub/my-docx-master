# =============================================================================
# 以 SAQ_scores_raw 为基础（235 名患者），从 follow_function_oct_filled 取数据
#
# 规则：
#   - 以 SAQ_scores_raw 为基准，保留全部行（左连接）
#   - 唯一匹配条件：patient_name（姓名，去空格后精确匹配）
#   - 源表中 SAQ 没有的列  -> 作为新列【追加】
#   - SAQ 里已有但为空的这几列（patient_id/procedure_date/sex/age）-> 从源表【回填】
#     （这几列在 SAQ 原表就是空列，若不特殊处理会一直是 NA）
#   - 某姓名在源表中找不到 -> 相关（追加+回填）列用 NA 占位，并标【棕色】
#
# 注意：姓名可能重名。源表同名多行时默认保留第一条并打印警告（可能取错人）。
#       整表全新写出（不用 loadWorkbook），因此颜色标注可靠、不会丢失。
#
# 输出：SAQ_output/SAQ_with_follow.xlsx
# =============================================================================

# ---- 1. 依赖包（首次运行请先取消下一行注释安装）-----------------------------
# install.packages(c("openxlsx", "dplyr"))
library(openxlsx)
library(dplyr)

# ---- 2. 路径与参数（如列名或文件名不同，只改这里）---------------------------
data_dir   <- "D:/R_project/OCT研究/SAQ_data"
output_dir <- "D:/R_project/OCT研究/SAQ_output"

base_file   <- "SAQ_scores_raw.xlsx"                # 基准表（235 名患者）
source_file <- "follow_function_oct_filled.xlsx"    # 数据源
out_file    <- "SAQ_with_follow.xlsx"

name_col <- "patient_name"     # 匹配用姓名列名（两张表须一致）

# SAQ 里已有但为空、需要从源表回填的列（如还有别的空列，加进来即可）
fill_cols <- c("patient_id", "procedure_date", "sex", "age")

brown_fill <- "#B5651D"        # 未匹配（NA 占位）标注色（棕色，可改）

# ---- 3. 读取（用 openxlsx：识别日期、"NA"当空、保留空列）--------------------
base <- read.xlsx(file.path(data_dir, base_file),   sheet = 1, detectDates = TRUE,
                  na.strings = "NA", skipEmptyCols = FALSE, skipEmptyRows = FALSE)
src  <- read.xlsx(file.path(data_dir, source_file), sheet = 1, detectDates = TRUE,
                  na.strings = "NA", skipEmptyCols = FALSE, skipEmptyRows = FALSE)

# ---- 4. 姓名归一化（去所有空格，含全角空格；不改大小写）--------------------
norm_name <- function(x) {
  x <- as.character(x)
  x <- gsub("[[:space:]]+", "", x)
  x <- gsub("　", "", x, fixed = TRUE)
  x
}
base$.name <- norm_name(base[[name_col]])
src$.name  <- norm_name(src[[name_col]])

# ---- 5. 源表按姓名去重（去空姓名；每姓名留第一条）--------------------------
src <- src[!is.na(src$.name) & src$.name != "", , drop = FALSE]
dups <- sum(duplicated(src$.name))
if (dups > 0)
  message(sprintf("注意：源表中有 %d 个重复 patient_name，已保留每姓名第一条（姓名匹配可能不唯一，请核对）。", dups))
src <- src[!duplicated(src$.name), ]

# ---- 6. 分列：追加列 / 回填列 -----------------------------------------------
add_cols <- setdiff(names(src), c(names(base), ".name"))                 # 源表独有 -> 追加
fill_now <- intersect(fill_cols, intersect(names(base), names(src)))     # 两表都有 -> 回填
if (length(setdiff(fill_cols, names(base))) > 0)
  message("提示：SAQ 缺少列（无法回填）：", paste(setdiff(fill_cols, names(base)), collapse = ", "))
if (length(setdiff(fill_cols, names(src)))  > 0)
  message("提示：源表缺少列（无法回填）：", paste(setdiff(fill_cols, names(src)),  collapse = ", "))

# 查找表：.name + 追加列(原名) + 回填列(改名 __f) + 匹配标记
lu <- src[, ".name", drop = FALSE]
for (col in add_cols) lu[[col]]              <- src[[col]]
for (col in fill_now) lu[[paste0(col, "__f")]] <- src[[col]]
lu$.matched <- TRUE

# ---- 7. 左连接 + 回填（只填 base 的 NA；整列全空则直接用源值以保留类型）------
merged  <- base %>% left_join(lu, by = ".name")
matched <- !is.na(merged$.matched)

for (col in fill_now) {
  fv <- merged[[paste0(col, "__f")]]
  if (all(is.na(merged[[col]]))) {
    merged[[col]] <- fv
  } else {
    need <- is.na(merged[[col]]) & !is.na(fv)
    merged[[col]][need] <- fv[need]
  }
}

# ---- 8. 整理输出（去辅助列；4 列在原位置，其余源列追加在后）------------------
out <- merged %>% select(-any_of(c(".name", ".matched", paste0(fill_now, "__f"))))
color_cols <- c(fill_now, add_cols)   # 未匹配时需标棕色的列

# ---- 9. 汇总信息 -------------------------------------------------------------
cat("SAQ 行数：          ", nrow(base), "\n")
cat("成功匹配：          ", sum(matched), "\n")
cat("未匹配（棕色标注）：", sum(!matched), "\n")
cat("已回填的列：        ", paste(fill_now, collapse = ", "), "\n")
cat("追加的列数：        ", length(add_cols), "  输出总列数：", ncol(out), "\n")

# ---- 10. 用 openxlsx 全新写出并着色 -----------------------------------------
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

wb <- createWorkbook()
addWorksheet(wb, "SAQ_follow")
writeData(wb, "SAQ_follow", out, keepNA = TRUE, na.string = "NA")   # NA 显示为 "NA" 占位

brownStyle <- createStyle(fgFill = brown_fill)
# 棕色：未匹配行中、追加列与回填列里仍为 NA 的单元格（逐列应用）
for (col in color_cols) {
  ci <- match(col, names(out))
  if (is.na(ci)) next
  rows_brown <- which(!matched & is.na(out[[col]]))
  if (length(rows_brown) > 0)
    addStyle(wb, "SAQ_follow", brownStyle, rows = rows_brown + 1, cols = ci,   # +1 跳过表头
             gridExpand = TRUE, stack = TRUE)
}

out_path <- file.path(output_dir, out_file)
saveWorkbook(wb, out_path, overwrite = TRUE)
cat("已生成：", out_path, "\n")
