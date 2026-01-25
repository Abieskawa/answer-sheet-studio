#!/usr/bin/env Rscript

ensure_packages <- function(pkgs) {
  missing <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]
  if (length(missing) == 0) {
    return(invisible(TRUE))
  }
  message("Installing missing R packages: ", paste(missing, collapse = ", "))
  tryCatch(
    {
      install.packages(missing, repos = "https://cloud.r-project.org")
      invisible(TRUE)
    },
    error = function(e) {
      stop(
        "Missing R packages and auto-install failed: ",
        paste(missing, collapse = ", "),
        "\nError: ",
        conditionMessage(e)
      )
    }
  )
}

required_pkgs <- c("readr", "dplyr", "tidyr", "ggplot2")
ensure_packages(required_pkgs)

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
})

theme_answer_sheet <- function(base_size = 13) {
  theme_minimal(base_size = base_size) +
    theme(
      plot.title = element_text(face = "bold", size = rel(1.15), margin = margin(b = 6)),
      plot.subtitle = element_text(size = rel(0.95), color = "#4b5563", margin = margin(b = 10)),
      plot.caption = element_text(size = rel(0.85), color = "#6b7280", margin = margin(t = 10)),
      axis.title.x = element_text(margin = margin(t = 8)),
      axis.title.y = element_text(margin = margin(r = 8)),
      panel.grid.minor = element_blank(),
      panel.grid.major.x = element_blank(),
      panel.grid.major.y = element_line(color = "#e5e7eb", linewidth = 0.6),
      legend.title = element_text(face = "bold"),
      legend.position = "top",
      strip.text = element_text(face = "bold"),
      strip.background = element_rect(fill = "#f3f4f6", color = NA),
      plot.background = element_rect(fill = "white", color = NA)
    )
}

get_arg <- function(args, key, default = NULL) {
  idx <- which(args == key)
  if (length(idx) == 0) {
    return(default)
  }
  if (idx[1] >= length(args)) {
    return(default)
  }
  args[[idx[1] + 1]]
}

args <- commandArgs(trailingOnly = TRUE)
input_path <- get_arg(args, "--input")
outdir <- get_arg(args, "--outdir")
lang <- get_arg(args, "--lang", default = "")

if (is.null(input_path) || input_path == "") {
  stop("Missing required --input <template.csv>")
}
if (is.null(outdir) || outdir == "") {
  stop("Missing required --outdir <output_dir>")
}

dir.create(outdir, showWarnings = FALSE, recursive = TRUE)

lang_norm <- tolower(gsub("-", "_", trimws(as.character(lang))))
is_zh <- startsWith(lang_norm, "zh")

tr <- function(zh, en) {
  if (is_zh) zh else en
}

normalize_answer <- function(x) {
  x <- toupper(trimws(as.character(x)))
  x[x == ""] <- NA
  x
}

df <- read_csv(
  input_path,
  show_col_types = FALSE,
  na = c("", "NA")
)

required_cols <- c("number", "correct", "points")
missing_cols <- setdiff(required_cols, names(df))
if (length(missing_cols) > 0) {
  stop(paste0("Missing columns: ", paste(missing_cols, collapse = ", ")))
}

df <- df %>%
  mutate(
    number = suppressWarnings(as.integer(number)),
    correct = normalize_answer(correct),
    points = suppressWarnings(as.numeric(points))
  )

student_cols <- setdiff(names(df), required_cols)
if (length(student_cols) < 1) {
  stop("No student columns found (expected columns after number/correct/points).")
}

answers <- df %>%
  select(all_of(student_cols)) %>%
  mutate(across(everything(), normalize_answer))

q_count <- nrow(df)
s_count <- length(student_cols)

correct_vec <- df$correct
points_vec <- df$points
points_vec[is.na(points_vec)] <- 0

ans_mat <- as.matrix(answers)
key_mat <- matrix(correct_vec, nrow = q_count, ncol = s_count, byrow = FALSE)
pt_mat <- matrix(points_vec, nrow = q_count, ncol = s_count, byrow = FALSE)

is_correct <- ans_mat == key_mat
is_correct[is.na(is_correct)] <- FALSE
is_correct[is.na(key_mat)] <- FALSE

score_mat <- is_correct * pt_mat
scores <- colSums(score_mat, na.rm = TRUE)
blank_counts <- colSums(is.na(ans_mat))
total_possible <- sum(points_vec[!is.na(correct_vec)], na.rm = TRUE)

scores_df <- tibble(
  person_id = student_cols,
  score = round(as.numeric(scores), 2),
  blank_count = as.integer(blank_counts),
  total_possible = round(as.numeric(total_possible), 2),
  percent = ifelse(total_possible > 0, round(score / total_possible, 2), NA_real_)
) %>%
  arrange(desc(score), person_id)

write_excel_csv(scores_df, file.path(outdir, "analysis_scores.csv"))

choices <- c("A", "B", "C", "D", "E")
has_key <- !is.na(correct_vec)

row_mean_bool <- function(m) {
  if (is.null(dim(m))) {
    return(as.numeric(m))
  }
  out <- rowMeans(m)
  as.numeric(out)
}

difficulty <- row_mean_bool(is_correct)
difficulty[!has_key] <- NA

student_order <- order(scores, student_cols)
group_n <- if (s_count > 30) as.integer(floor(s_count * 0.27)) else as.integer(floor(s_count / 2))
low_idx <- if (group_n > 0) student_order[seq_len(group_n)] else integer(0)
high_idx <- if (group_n > 0) student_order[(s_count - group_n + 1):s_count] else integer(0)

disc <- rep(NA_real_, q_count)
if (length(low_idx) > 0 && length(high_idx) > 0) {
  high_diff <- row_mean_bool(is_correct[, high_idx, drop = FALSE])
  low_diff <- row_mean_bool(is_correct[, low_idx, drop = FALSE])
  disc <- high_diff - low_diff
  disc[!has_key] <- NA
}

blank_rate <- rowSums(is.na(ans_mat)) / s_count
multi_rate <- rowSums(!is.na(ans_mat) & nchar(ans_mat) > 1) / s_count
choice_rates <- lapply(
  choices,
  function(ch) {
    rowSums(ans_mat == ch, na.rm = TRUE) / s_count
  }
)
names(choice_rates) <- paste0("p_", choices)

recognized_choices <- Reduce(`|`, lapply(choices, function(ch) ans_mat == ch))
recognized_choices[is.na(recognized_choices)] <- FALSE
other_rate <- rowSums(!is.na(ans_mat) & !recognized_choices & nchar(ans_mat) <= 1) / s_count

item_df <- tibble(
  number = df$number,
  correct = correct_vec,
  points = points_vec,
  difficulty = ifelse(is.na(difficulty), NA_real_, round(difficulty, 2)),
  discrimination = ifelse(is.na(disc), NA_real_, round(disc, 2)),
  blank_rate = round(blank_rate, 2),
  multi_rate = round(multi_rate, 2),
  other_rate = round(other_rate, 2)
)
for (nm in names(choice_rates)) {
  item_df[[nm]] <- round(choice_rates[[nm]], 2)
}

write_excel_csv(item_df, file.path(outdir, "analysis_item.csv"))

score_vec <- scores_df$score
qv <- as.numeric(quantile(score_vec, c(0.88, 0.75, 0.5, 0.25, 0.12), na.rm = TRUE))
q88 <- qv[1]
q75 <- qv[2]
q50 <- qv[3]
q25 <- qv[4]
q12 <- qv[5]

summary_df <- tibble(
  students = s_count,
  questions = q_count,
  total_possible = total_possible,
  mean = round(mean(score_vec, na.rm = TRUE), 2),
  sd = round(sd(score_vec, na.rm = TRUE), 2),
  p88 = round(q88, 2),
  p75 = round(q75, 2),
  median = round(q50, 2),
  p25 = round(q25, 2),
  p12 = round(q12, 2)
)
write_excel_csv(summary_df, file.path(outdir, "analysis_summary.csv"))

binwidth <- 1
if (!is.na(total_possible) && total_possible > 0) {
  binwidth <- max(1, round(total_possible / 20))
}

metric_lines <- tibble(
  x = c(mean(score_vec, na.rm = TRUE), q50, q88, q75, q25, q12),
  label = c(tr("平均", "Mean"), tr("中位數", "Median"), "P88", "P75", "P25", "P12"),
  group = c(tr("平均", "Mean"), tr("中位數", "Median"), tr("百分位數", "Percentiles"), tr("百分位數", "Percentiles"), tr("百分位數", "Percentiles"), tr("百分位數", "Percentiles"))
)

max_count <- max(as.integer(table(floor(score_vec / binwidth) * binwidth)), 1, na.rm = TRUE)
line_label_df <- metric_lines %>%
  mutate(
    y = max_count + 0.35,
    x = pmax(0, x)
  )

mean_label <- tr("平均", "Mean")
median_label <- tr("中位數", "Median")
percentiles_label <- tr("百分位數", "Percentiles")

p_hist <- ggplot(scores_df, aes(x = score)) +
  geom_histogram(
    binwidth = binwidth,
    boundary = 0,
    fill = "#0ea5e9",
    color = "white",
    alpha = 0.92
  ) +
  geom_vline(
    data = metric_lines,
    aes(xintercept = x, color = group, linetype = label),
    linewidth = 1.1,
    show.legend = FALSE
  ) +
  geom_label(
    data = line_label_df,
    aes(x = x, y = y, label = label, color = group),
    size = 3.2,
    linewidth = 0,
    fill = "white",
    alpha = 0.85,
    show.legend = FALSE
  ) +
  scale_color_manual(
    name = tr("指標線", "Reference lines"),
    values = setNames(c("#ef4444", "#2563eb", "#6b7280"), c(mean_label, median_label, percentiles_label))
  ) +
  scale_linetype_manual(
    name = tr("指標線", "Reference lines"),
    values = setNames(
      c("solid", "solid", "dashed", "dashed", "dashed", "dashed"),
      c(mean_label, median_label, "P88", "P75", "P25", "P12")
    )
  ) +
  scale_x_continuous(breaks = function(l) pretty(l, n = 10)) +
  scale_y_continuous(
    breaks = function(l) seq(0, ceiling(max(l, na.rm = TRUE)), by = 1),
    expand = expansion(mult = c(0, 0.12))
  ) +
  theme_answer_sheet(base_size = 14) +
  labs(
    title = tr("成績分佈", "Score distribution"),
    subtitle = if (is_zh) {
      sprintf("學生數：%d｜滿分：%s｜分數級距：%s", s_count, format(total_possible, trim = TRUE), format(binwidth, trim = TRUE))
    } else {
      sprintf("Students: %d | Full score: %s | Binwidth: %s", s_count, format(total_possible, trim = TRUE), format(binwidth, trim = TRUE))
    },
    x = tr("分數", "Score"),
    y = tr("學生人數", "Students")
  )

ggsave(
  filename = file.path(outdir, "analysis_score_hist.png"),
  plot = p_hist,
  width = 8,
  height = 4.5,
  dpi = 200,
  bg = "white"
)

item_plot_df <- item_df %>%
  select(number, difficulty, discrimination) %>%
  pivot_longer(cols = c(difficulty, discrimination), names_to = "metric", values_to = "value") %>%
  mutate(metric = recode(metric, difficulty = tr("難度（正答率）", "Difficulty (accuracy)"), discrimination = tr("鑑別度", "Discrimination")))

ref_lines <- tibble(
  metric = c(rep(tr("難度（正答率）", "Difficulty (accuracy)"), 3), rep(tr("鑑別度", "Discrimination"), 3)),
  y = c(0.25, 0.5, 0.75, 0.0, 0.2, 0.4),
  label = c("0.25", "0.50", "0.75", "0.00", "0.20", "0.40")
)

p_item <- ggplot(item_plot_df, aes(x = number, y = value)) +
  geom_hline(
    data = ref_lines,
    aes(yintercept = y),
    linewidth = 0.6,
    color = "#d1d5db",
    linetype = "dashed",
    inherit.aes = FALSE
  ) +
  geom_label(
    data = ref_lines,
    aes(x = Inf, y = y, label = label),
    inherit.aes = FALSE,
    hjust = 1.08,
    vjust = -0.45,
    size = 3.0,
    linewidth = 0,
    label.size = 0,
    fill = "white",
    alpha = 0.72,
    color = "#6b7280"
  ) +
  geom_line(color = "#94a3b8", linewidth = 0.6, na.rm = TRUE) +
  geom_point(size = 1.9, color = "#0ea5e9", alpha = 0.95, na.rm = TRUE) +
  facet_wrap(~metric, ncol = 1, scales = "free_y") +
  scale_y_continuous(labels = function(x) sprintf("%.2f", x)) +
  scale_x_continuous(breaks = function(l) pretty(l, n = 12)) +
  coord_cartesian(clip = "off") +
  theme_answer_sheet(base_size = 13) +
  labs(
    title = tr("題目分析", "Item analysis"),
    subtitle = tr(
      "難度與鑑別度（虛線：難度 0.25/0.50/0.75；鑑別度 0.00/0.20/0.40）",
      "Difficulty and discrimination (dashed lines: difficulty 0.25/0.50/0.75; discrimination 0.00/0.20/0.40)"
    ),
    x = tr("題號", "Question"),
    y = NULL
  ) +
  theme(
    legend.position = "none",
    plot.margin = margin(5.5, 24, 5.5, 5.5)
  )

ggsave(
  filename = file.path(outdir, "analysis_item_plot.png"),
  plot = p_item,
  width = 8,
  height = 7,
  dpi = 200,
  bg = "white"
)
