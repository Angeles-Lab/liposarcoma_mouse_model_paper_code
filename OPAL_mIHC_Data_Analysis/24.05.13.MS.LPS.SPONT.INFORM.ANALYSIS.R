# Created by phenoptr 0.3.2 and phenoptrReports 0.3.3 on 2024-05-13
# http://akoyabio.github.io/phenoptr
# http://akoyabio.github.io/phenoptrReports

suppressPackageStartupMessages(library(dplyr))
library(phenoptr)
library(phenoptrReports)
library(openxlsx)

# Read the consolidated data file
csd_path =
  "C:/Users/ekenna/OneDrive - Michigan Medicine/Desktop/UMICH/0 - ANGELES LAB/R STUDIO/PHENOPTRREPORTS/Consolidated_data.txt"
csd = read_cell_seg_data(csd_path, col_select="phenoptrReports")

# Make a table summarizing the number of fields per slide
summary_table = csd %>%
  group_by(`Slide ID`) %>%
  summarize(`Number of fields`=n_distinct(`Annotation ID`))

tissue_categories = c("All")

# Define phenotypes
phenotypes = parse_phenotypes("CD3+/FOXP3+", "CD3+", "CD3+/CD8a+", "CD3+/CD8a+/PD1+", "CD19+", "F480+", "Total Cells")

# Column to aggregate by
.by = "Annotation ID"

# Count phenotypes per tissue category
counts = count_phenotypes(csd, phenotypes, tissue_categories, .by=.by)
percents = counts_to_percents(counts)

expression_params = NULL

# Summarize nearest neighbor distances
nearest_detail_path = file.path(
  "C:/Users/ekenna/OneDrive - Michigan Medicine/Desktop/UMICH/0 - ANGELES LAB/R STUDIO/PHENOPTRREPORTS/ANALYSIS",
  "nearest_neighbors.txt")
nearest_neighbors = nearest_neighbor_summary(
  csd, phenotypes, tissue_categories, nearest_detail_path, .by=.by,
  extra_cols=expression_params)

# Summary of cells within a specific distance
radii = 15
count_detail_path = file.path(
  "C:/Users/ekenna/OneDrive - Michigan Medicine/Desktop/UMICH/0 - ANGELES LAB/R STUDIO/PHENOPTRREPORTS/ANALYSIS",
  "count_within.txt")
count_within = count_within_summary(
  csd, radii, phenotypes, tissue_categories,
  count_detail_path, .by=.by, extra_cols=expression_params)

# Clean up the aggregation column
# Do this at the end or it will break merges
cleanup = function(d) {
  by_col = ifelse(.by %in% names(d), .by, 'Slide ID')
  d[[by_col]] = stringr::str_remove_all(d[[by_col]], '^N')
  d
}

summary_table = cleanup(summary_table)
counts = cleanup(counts)
percents = cleanup(percents)
nearest_neighbors = cleanup(nearest_neighbors)
count_within = cleanup(count_within)
 
# Write it all out to an Excel workbook
wb = createWorkbook()
write_summary_sheet(wb, summary_table)
write_counts_sheet(wb, counts)
write_percents_sheet(wb, percents)
write_nearest_neighbor_summary_sheet(wb, nearest_neighbors)
write_count_within_sheet(wb, count_within)

workbook_path = file.path(
  "C:/Users/ekenna/OneDrive - Michigan Medicine/Desktop/UMICH/0 - ANGELES LAB/R STUDIO/PHENOPTRREPORTS/ANALYSIS",
  "Results.xlsx")
if (file.exists(workbook_path)) file.remove(workbook_path)
saveWorkbook(wb, workbook_path)

# Write summary charts
charts_path = file.path(
  "C:/Users/ekenna/OneDrive - Michigan Medicine/Desktop/UMICH/0 - ANGELES LAB/R STUDIO/PHENOPTRREPORTS/ANALYSIS",
  "Charts.docx")
if (file.exists(charts_path)) file.remove(charts_path)
write_summary_charts(workbook_path, charts_path, .by=.by)

# Save session info
info_path = file.path(
  "C:/Users/ekenna/OneDrive - Michigan Medicine/Desktop/UMICH/0 - ANGELES LAB/R STUDIO/PHENOPTRREPORTS/ANALYSIS",
  "session_info.txt")
write_session_info(info_path)
