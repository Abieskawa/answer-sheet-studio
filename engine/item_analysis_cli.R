#!/usr/bin/env Rscript

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
})

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

if (is.null(input_path) || input_path == "") {
  stop("Missing required --input <template.csv>")
}
if (is.null(outdir) || outdir == "") {
  stop("Missing required --outdir <output_dir>")
}

dir.create(outdir, showWarnings = FALSE, recursive = TRUE)

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
  score = as.numeric(scores),
  blank_count = as.integer(blank_counts),
  total_possible = as.numeric(total_possible),
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
half <- as.integer(floor(s_count / 2))
low_idx <- if (half > 0) student_order[seq_len(half)] else integer(0)
high_idx <- if (half > 0) student_order[(s_count - half + 1):s_count] else integer(0)

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

summary_df <- tibble(
  students = s_count,
  questions = q_count,
  total_possible = total_possible,
  mean = round(mean(scores_df$score), 3),
  sd = round(sd(scores_df$score), 3),
  p88 = round(quantile(scores_df$score, 0.88, na.rm = TRUE), 3),
  p75 = round(quantile(scores_df$score, 0.75, na.rm = TRUE), 3),
  median = round(quantile(scores_df$score, 0.5, na.rm = TRUE), 3),
  p25 = round(quantile(scores_df$score, 0.25, na.rm = TRUE), 3),
  p12 = round(quantile(scores_df$score, 0.12, na.rm = TRUE), 3)
)
write_excel_csv(summary_df, file.path(outdir, "analysis_summary.csv"))

binwidth <- 1
if (!is.na(total_possible) && total_possible > 0) {
  binwidth <- max(1, round(total_possible / 20))
}

p_hist <- ggplot(scores_df, aes(x = score)) +
  geom_histogram(binwidth = binwidth, fill = "#0ea5e9", color = "white", alpha = 0.9) +
  geom_vline(xintercept = mean(scores_df$score), color = "#ef4444", linewidth = 1) +
  geom_vline(xintercept = median(scores_df$score), color = "#3b82f6", linewidth = 1) +
  theme_minimal(base_size = 14) +
  labs(title = "Score distribution", x = "Score", y = "Students") +
  theme(plot.title = element_text(face = "bold"))

ggsave(
  filename = file.path(outdir, "analysis_score_hist.png"),
  plot = p_hist,
  width = 8,
  height = 4.5,
  dpi = 160
)

item_plot_df <- item_df %>%
  select(number, difficulty, discrimination, blank_rate) %>%
  pivot_longer(cols = c(difficulty, discrimination, blank_rate), names_to = "metric", values_to = "value")

p_item <- ggplot(item_plot_df, aes(x = number, y = value)) +
  geom_line(color = "#e5e7eb") +
  geom_point(size = 1.6, color = "#e5e7eb") +
  facet_wrap(~metric, ncol = 1, scales = "free_y") +
  theme_minimal(base_size = 13) +
  labs(title = "Item metrics", x = "Question", y = NULL) +
  theme(
    plot.title = element_text(face = "bold"),
    strip.text = element_text(face = "bold")
  )

ggsave(
  filename = file.path(outdir, "analysis_item_plot.png"),
  plot = p_item,
  width = 8,
  height = 7,
  dpi = 160
)
