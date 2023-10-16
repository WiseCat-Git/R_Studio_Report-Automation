# Load required libraries
library(officer)
library(flextable)
library(ggplot2)
library(dplyr)

# Load the Excel file (replace 'your_excel_file.xlsx' with your file path)
excel_data <- readxl::read_excel("your_excel_file.xlsx")

# Create a PowerPoint document
ppt <- read_pptx()

# Add a slide with a title and content layout
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")

# Add a title to the slide
ppt <- ph_with_text(ppt, str = "My PowerPoint Report", location = ph_location_type(type = "title"))

# Add a subtitle with the last update date (replace 'last_update_date' with your date)
last_update_date <- "2023-10-05"  # Replace with your actual date
subtitle <- paste("Last Updated:", last_update_date)
ppt <- ph_with_text(ppt, str = subtitle, location = ph_location_type(type = "subtitle"))

# Create a Flextable from the Excel data
excel_table <- flextable::qflextable(excel_data)

# Customize table formatting (optional)
excel_table <- flextable::set_table_properties(excel_table, width = .8)

# Insert the Excel table into the PowerPoint slide
ppt <- ph_with_flextable(ppt, value = excel_table, location = ph_location_type(type = "body"))

# Create a sample bar chart (replace 'your_data' and 'your_x'/'your_y' with actual data)
sample_data <- data.frame(
  Category = c("A", "B", "C", "D"),
  Value = c(25, 50, 30, 45)
)

bar_chart <- ggplot(sample_data, aes(x = Category, y = Value)) +
  geom_bar(stat = "identity", fill = "skyblue") +
  labs(title = "Sample Bar Chart", x = "Categories", y = "Values")

# Insert the bar chart into the PowerPoint slide
ppt <- ph_with_gg(ppt, code = bar_chart, location = ph_location_type(type = "body"))

# Add a text box with insights
insights_text <- "Here are some insights based on the data:\n\n1. Insight 1\n2. Insight 2\n3. Insight 3"

ppt <- ph_with_text(ppt, str = insights_text, location = ph_location_type(type = "body"))

# Save the PowerPoint file
save_pptx(ppt, "my_powerpoint_report.pptx")

# View the generated PowerPoint
file.edit("my_powerpoint_report.pptx")
