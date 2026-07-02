# =============================================================================
# 三个 Excel 文件按 patient_id + procedure_date 匹配合并
#
# 匹配规则：
#   - patient_id 与 procedure_date 同时相同才算匹配（三张表都存在）
#   - patient_id 前导 0 视为同一 ID（0054288752 == 54288752）
#   - procedure_date 统一转成 YYYY-MM-DD 后再比较
#
# 输出：
#   - 对匹配上的病例，提取三张表的全部列
#   - 列顺序：OCT_raw_data -> function_raw_20260507_1 -> follow_data
#   - patient_id / procedure_date 只保留一份，放在最前面
# =============================================================================

# ---- 1. 依赖包（首次运行请先取消下一行注释安装）-----------------------------
# install.packages(c("readxl", "writexl", "dplyr"))
library(readxl)
library(writexl)
library(dplyr)

# ---- 2. 路径与参数（如列名或文件名不同，只改这里）---------------------------
data_dir   <- "D:/R_project/随访+OCT+功能/data"
output_dir <- "D:/R_project/随访+OCT+功能/output"

oct_file    <- "OCT_raw_data.xlsx"
func_file   <- "function_raw_20260507_1.xlsx"
follow_file <- "follow_data.xlsx"

id_col   <- "patient_id"       # 病例 ID 列名（三张表须一致）
date_col <- "procedure_date"   # 手术/检查日期列名（三张表须一致）

out_file <- "matched_data.xlsx"

# ---- 3. 读取三张表 -----------------------------------------------------------
oct    <- read_excel(file.path(data_dir, oct_file))
func   <- read_excel(file.path(data_dir, func_file))
follow <- read_excel(file.path(data_dir, follow_file))

# ---- 4. 归一化函数 -----------------------------------------------------------
# patient_id：去掉前导 0，使 0054288752 与 54288752 视为同一 ID
norm_id <- function(x) {
  na <- is.na(x)
  if (is.numeric(x)) {
    x <- sprintf("%.0f", x)      # 数值型：避免科学计数法，保留完整整数
  } else {
    x <- as.character(x)
  }
  x <- trimws(x)
  x <- sub("^0+", "", x)         # 去掉前导 0
  x[x == ""] <- "0"              # 防止全 0 被清空
  x[na] <- NA
  x
}

# procedure_date：统一成 "YYYY-MM-DD" 字符串，兼容日期型 / Excel序列号 / 文本
norm_date <- function(x) {
  if (inherits(x, "Date"))   return(format(x, "%Y-%m-%d"))
  if (inherits(x, "POSIXt")) return(format(as.Date(x), "%Y-%m-%d"))
  if (is.numeric(x))         return(format(as.Date(x, origin = "1899-12-30"), "%Y-%m-%d"))
  x <- trimws(as.character(x))
  d <- as.Date(x, tryFormats = c("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y",
                                 "%d/%m/%Y", "%Y.%m.%d", "%Y年%m月%d日"))
  format(d, "%Y-%m-%d")
}

# ---- 5. 生成匹配键 -----------------------------------------------------------
oct    <- oct    %>% mutate(.pid = norm_id(.data[[id_col]]),   .pdate = norm_date(.data[[date_col]]))
func   <- func   %>% mutate(.pid = norm_id(.data[[id_col]]),   .pdate = norm_date(.data[[date_col]]))
follow <- follow %>% mutate(.pid = norm_id(.data[[id_col]]),   .pdate = norm_date(.data[[date_col]]))

# 提示：若同一张表内存在重复的 (patient_id, procedure_date)，合并时对应行会成倍展开
dup_note <- function(df, name) {
  k <- paste(df$.pid, df$.pdate)
  d <- sum(duplicated(k))
  if (d > 0) message(sprintf("注意：%s 存在 %d 条重复的 patient_id+procedure_date，匹配后相关行会展开。", name, d))
}
dup_note(oct, oct_file); dup_note(func, func_file); dup_note(follow, follow_file)

# ---- 6. 三表内连接（列顺序 OCT -> function -> follow）------------------------
# function / follow 表去掉重复的 id、date 列，只保留一份（来自 OCT）
func_j   <- func   %>% select(-all_of(c(id_col, date_col)))
follow_j <- follow %>% select(-all_of(c(id_col, date_col)))

merged <- oct %>%
  inner_join(func_j,   by = c(".pid", ".pdate"), suffix = c("", "_function")) %>%
  inner_join(follow_j, by = c(".pid", ".pdate"), suffix = c("", "_follow"))

# 去掉辅助键，并把 patient_id / procedure_date 放到最前面
merged <- merged %>%
  select(-.pid, -.pdate) %>%
  relocate(all_of(c(id_col, date_col)))

# ---- 7. 打印汇总信息 ---------------------------------------------------------
cat("OCT_raw_data      行数：", nrow(oct),    "\n")
cat("function_raw      行数：", nrow(func),   "\n")
cat("follow_data       行数：", nrow(follow), "\n")
cat("匹配上的病例行数： ",       nrow(merged), "\n")
cat("输出列数：         ",       ncol(merged), "\n")

# ---- 8. 写出结果 -------------------------------------------------------------
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
out_path <- file.path(output_dir, out_file)
write_xlsx(merged, out_path)
cat("已生成：", out_path, "\n")
