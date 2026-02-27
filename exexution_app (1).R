library(shiny)
library(bslib)
library(DT)
library(tidyverse)
library(visNetwork)
library(openxlsx)

open_orders = read.xlsx("data.xlsx",sheet="Current Open Orders", detectDates = TRUE)

historical_orders = read.xlsx("data.xlsx",sheet="Historical Orders", detectDates = TRUE)

fg_supply = read.xlsx("data.xlsx",sheet="FG_Projected_supply", detectDates = TRUE)

fg_allocation = read.xlsx("data.xlsx",sheet="FG_allocation", detectDates = TRUE)

production_schedule = read.xlsx("data.xlsx",sheet="Daily Production Schedule", detectDates = TRUE)

bom = read.xlsx("data.xlsx",sheet="BOM", detectDates = TRUE)

rm_projected_supply = read.xlsx("data.xlsx",sheet="RM_projected_supply", detectDates = TRUE)

sourcing_parameters = read.xlsx("data.xlsx",sheet="sourcing_parameters", detectDates = TRUE)

historical_orders <- historical_orders %>%
  group_by(customer_code) %>%
  summarise(Average = mean(quantity, na.rm = TRUE) + 2*(sd(quantity, na.rm = TRUE))) %>%
  ungroup()

bulk <- open_orders %>%
  left_join(historical_orders, by = "customer_code") %>%
  mutate(
    status = case_when(
      quantity > Average ~ "Bulk",
      quantity == Average ~ "Regular",
      quantity < Average ~ "Regular",
      TRUE ~ "No data"
    )) %>%
  select (customer_code,customer_desc,quantity,Average,status)


fg_allocation_filter <- fg_allocation %>%
  select(customer_code,customer_desc,`2023-04-17`)

total_available_stock = fg_supply %>%
  filter(type == "total_Fgsupply (Opening Inventory)") %>%
  select(`2023-04-17`) %>%
  pull()

confirm <- bulk %>%
  left_join(fg_allocation_filter, by="customer_code") %>%
  mutate(
    allocation_amount = `2023-04-17`* total_available_stock,
    confirmation_status = case_when(
      quantity > allocation_amount ~ "Partially confirmed",
      quantity < allocation_amount ~ "Fully confirmed"
    )) %>%
  select(customer_code,customer_desc.x,quantity,Average,status,`2023-04-17`,allocation_amount,confirmation_status)


# Bulk Order Early Availability Check

# Load the Data
# Production Schedule
production_schedule = production_schedule %>%
  pivot_longer(cols = c(3:23), names_to = "week", values_to = "production_plan")


# Projected Material Supply
rm_projected_supply = rm_projected_supply %>%
  # calculating the opening inventory of 14th April
  mutate(`2023-04-14` = `2023-04-13` - 3000 + `2023-04-14`) %>%
  select(-`2023-04-13`) %>%
  pivot_longer(
    cols = c(3:19),
    names_to = "week",
    values_to = "projected_rm_supply"
  )

## Data Transformation

order_review_date = as.Date("2023-04-10")

next_possible_production = if_else(
  wday(order_review_date + 4) == 7,
  order_review_date + 7,
  order_review_date + 4
)

# Decide the first date of additional production
additional_prod_day1 = next_possible_production

additional_prod_day2 = if_else(
  wday(next_possible_production + 1) == 7,
  next_possible_production + 3,
  next_possible_production
)

# Decide the daily additional quantities
additional_prod_day1_qty = 6000
additional_prod_day2_qty = 2000

# Calculate the daily material requirements
daily_material_requirement = bom %>%
  select(material, material_desc, required_quantity) %>%
  filter(required_quantity == 1, material_desc != "engine_module") %>%
  rename(component = "material", component_desc = "material_desc") %>%
  mutate(material = "Z22215") %>%
  left_join(production_schedule, by = "material") %>%
  filter(week > "2023-04-13") %>%
  mutate(
    addition_production = case_when(
      week == additional_prod_day1 ~ additional_prod_day1_qty,
      week == additional_prod_day2 ~ additional_prod_day2_qty,
      TRUE ~ 0
    )
  ) %>%
  left_join(rm_projected_supply, join_by(component == material, week)) %>%
  mutate(across(
    c(production_plan, projected_rm_supply),
    ~ replace_na(as.numeric(.), 0)
  )) %>%
  group_by(component) %>%
  mutate(
    daily_requirement = production_plan + addition_production,
    net_change = projected_rm_supply - daily_requirement,
    closing_inventory = round(cumsum(net_change), 0)
  )

# Show the output at material in rows and weeks in column
daily_rm_pivot = daily_material_requirement %>%
  select(component, component_desc, week, closing_inventory) %>%
  pivot_wider(names_from = week, values_from = closing_inventory)


datatable(daily_rm_pivot) %>%
  formatStyle(
    columns = 3:ncol(daily_rm_pivot), # All week columns (excluding component columns)
    backgroundColor = styleInterval(
      cuts = c(0),
      values = c('#ffcccc', '#ccffcc') # red for <0, green for >=0
    )
  )


# ======================================================
# TASK 5 REVISED: Scheduling AFTER Current Production Run
# ======================================================

# --- STEP 1: CALCULATE THE "PRODUCTION FINISH DATE" ---

# 1.1 Identify Missing Parts (Same as before)
production_days <- c(additional_prod_day1, additional_prod_day2)
missing_parts <- daily_material_requirement %>%
  filter(week %in% production_days, closing_inventory < 0) %>%
  distinct(component, component_desc)

# 1.2 Calculate Material Arrival Date (Same as before)
if(nrow(missing_parts) > 0) {
  recovery_plan <- missing_parts %>%
    left_join(sourcing_parameters, by = c("component" = "material")) %>%
    mutate(
      applied_lead_time = coalesce(lead_time_express, lead_time_normal), 
      part_arrival_date = order_review_date + applied_lead_time 
    )
  material_ready_date <- max(recovery_plan$part_arrival_date, na.rm = TRUE)
} else {
  material_ready_date <- order_review_date
}

# --- LOGIC FIX HERE ---
# We cannot start on April 14th because the line is booked.
# The next open week starts on Monday, April 17th.
next_open_slot <- as.Date("2023-04-17") 

# The Start Date is the LATER of:
# 1. Material Arrival
# 2. Next Open Production Slot (April 17th)
earliest_start <- max(next_open_slot, material_ready_date)

# Find Start Day (Skip Sat/Sun - Just in case date shifts)
new_prod_day1 <- earliest_start
while(wday(new_prod_day1) == 7 || wday(new_prod_day1) == 1) { new_prod_day1 <- new_prod_day1 + 1 }

# Find Finish Day (Day 1 + 1 Day)
new_prod_day2 <- new_prod_day1 + 1
while(wday(new_prod_day2) == 7 || wday(new_prod_day2) == 1) { new_prod_day2 <- new_prod_day2 + 1 }

final_production_date <- new_prod_day2

# --- STEP 2: BUILD FINAL TABLE (Same as before) ---
final_order_list <- confirm %>%
  select(customer_code, customer_desc = customer_desc.x, status, allocation_amount, confirmation_status) %>%
  left_join(open_orders, by = c("customer_code", "customer_desc")) %>%
  mutate(
    Confirmed_Stock_Qty = floor(pmin(quantity, allocation_amount)),
    Backlog_Qty = pmax(0, quantity - Confirmed_Stock_Qty),
    
    # Stock is available April 17th (Per Image 3)
    Stock_Supply_Date = as.Date("2023-04-17"),
    
    # Production is available on the NEW calculated date
    Backlog_Supply_Date = final_production_date
  )

final_output_table <- final_order_list %>%
  transmute(
    Order_ID = order_id,
    Customer = customer_desc,
    Type = status, 
    Final_Status = case_when(
      confirmation_status == "Fully confirmed" ~ "Filled from Stock",
      confirmation_status == "Partially confirmed" ~ "Split Shipment (Stock + Production)",
      status == "Bulk" ~ "New Production Run"
    ),
    Order_Qty = quantity,
    Stock_Fill = Confirmed_Stock_Qty,
    Stock_Date = Stock_Supply_Date,
    Production_Fill = Backlog_Qty,
    Production_Date = Backlog_Supply_Date,
    
    # Final Date Logic
    Final_Supply_Date = if_else(Backlog_Qty > 0, Production_Date, Stock_Date),
    Delay_Days = as.numeric(Final_Supply_Date - required_date),
    Required_Date = as.Date(required_date)
  )

# Display
datatable(final_output_table, 
          caption = "Final Order Confirmation (Corrected for Schedule Gap)",
          options = list(pageLength = 10, scrollX = TRUE)) %>%
  formatStyle(
    'Final_Status',
    backgroundColor = styleEqual(
      c("Filled from Stock", "Split Shipment (Stock + Production)", "New Production Run"), 
      c('#ccffcc', '#fff3cd', '#ffeb99')
    )
  ) %>%
  formatDate(c('Stock_Date', 'Production_Date', 'Final_Supply_Date'), method = 'toDateString')



# ======================================================
# TASK 6 (OPTIMIZED): Calculate Cost WITHOUT Night Shift Allowance
# ======================================================

# 1. Load Production Parameters
production_parameters = read.xlsx("data.xlsx", sheet="production_parameters", detectDates = TRUE)

# 2. Extract Individual Cost Components (To build a custom rate)
electricity_cost <- production_parameters %>%
  filter(item == "electricity") %>%
  pull(value) %>%
  as.numeric()

transport_cost <- production_parameters %>%
  filter(item == "transportation") %>%
  pull(value) %>%
  as.numeric()

food_cost <- production_parameters %>%
  filter(item == "food") %>%
  pull(value) %>%
  as.numeric()

# 3. Calculate "Optimized Day Rate" (Excluding 480k Allowance)
day_shift_cost <- electricity_cost + transport_cost + food_cost

# 4. Calculate Production Load
line_capacity <- production_parameters %>%
  filter(item == "Z22215_module_linecapacity") %>%
  pull(value) %>%
  as.numeric()

# Sum up the backlog from Task 5 output
total_additional_units <- sum(final_output_table$Production_Fill, na.rm = TRUE)
shifts_required <- ceiling(total_additional_units / line_capacity)

# 5. Calculate Total Savings
standard_rate <- production_parameters %>%
  filter(item == "Total operational_cost_pershift") %>%
  pull(value) %>%
  as.numeric()

standard_total_cost <- shifts_required * standard_rate
optimized_total_cost <- shifts_required * day_shift_cost
total_savings <- standard_total_cost - optimized_total_cost

# 6. Build the Financial Summary Table
task6_output <- data.frame(
  Metric = c(
    "Total Additional Units",
    "Shifts Required",
    "Standard Cost (Inc. Night Shift)",
    "Optimized Day-Shift Rate",
    "OPTIMIZED TOTAL COST",
    "TOTAL SAVINGS"
  ),
  Value = c(
    format(total_additional_units, big.mark = ","),
    as.character(shifts_required),
    paste(format(standard_total_cost, big.mark = ","), "INR"),
    paste(format(day_shift_cost, big.mark = ","), "INR"),
    paste(format(optimized_total_cost, big.mark = ","), "INR"),
    paste(format(total_savings, big.mark = ","), "INR")
  )
)

# 7. Display Output
datatable(task6_output, 
          caption = "Task 6: Financial Optimization (Avoiding Night Shift)",
          options = list(dom = 't', paging = FALSE),
          rownames = FALSE) %>%
  formatStyle(
    'Metric',
    target = 'row',
    backgroundColor = styleEqual(c("OPTIMIZED TOTAL COST", "TOTAL SAVINGS"), c('#ccffcc', '#ffeb99')),
    fontWeight = styleEqual(c("OPTIMIZED TOTAL COST", "TOTAL SAVINGS"), "bold")
  )

# ======================================================
# TASK 7: Final Order Confirmation (Client View)
# ======================================================

# Logic: 
# 1. We take the 'final_output_table' (which already joined open_orders in Task 5).
# 2. We select only the client-facing columns.
# 3. We assume "Confirmed Qty" is 100% of the Order Qty because we are fulfilling 
#    the backlog via the new production run (just at a later date).

task7_output <- final_output_table %>%
  select(
    Order_ID,
    Customer,
    "Requested Qty" = Order_Qty,
    "Confirmed Qty" = Order_Qty, 
    "Requested Date" = Required_Date,   # This column came from the join in Task 5
    "Confirmed Date" = Final_Supply_Date,
    "Delay (Days)" = Delay_Days
  )

# Display Task 7
datatable(task7_output, 
          caption = "Task 7: Final Order Confirmation Receipt",
          options = list(dom = 't', paging = FALSE),
          rownames = FALSE) %>%
  formatDate(c('Requested Date', 'Confirmed Date'), method = 'toDateString') %>%
  formatStyle(
    'Delay (Days)',
    color = styleInterval(0, c('green', 'red')), 
    fontWeight = 'bold'
  )


# ======================================================
# TASK 8: Financial Implications (Profit & Penalties)
# ======================================================

# Logic:
# 1. Revenue = Order Qty * Selling Price.
# 2. Penalties = Only applied if Delay_Days > 0.
#    (We calculate penalties ONLY on the 'Production_Fill' portion, as Stock fill was on time).
# 3. Net Profit = Gross Margin - Penalties - Operational Costs (from Task 6).

# 1. Load Financial Parameters
penalty_data = read.xlsx("data.xlsx", sheet="Penalty", detectDates = TRUE)

# 2. Calculate Financials per Order
financial_analysis <- final_output_table %>%
  # Join with Penalty sheet to get Price and Margin %
  left_join(penalty_data, by = c("Customer" = "customer_desc")) %>%
  mutate(
    # A. Revenue & Margin
    Total_Revenue = Order_Qty * Selling_price,
    Gross_Margin_Value = Total_Revenue * margin,
    
    # B. Penalty Calculation
    # If Delay > 0, Penalty = (Delayed Qty * Price * Penalty %)
    # Note: 'Production_Fill' is the quantity that was delayed (backlog)
    Penalty_Applicable = if_else(Delay_Days > 0, 1, 0),
    Penalty_Cost = Penalty_Applicable * (Production_Fill * Selling_price * Qty_Miss_Penalty)
  )

# 3. Aggregate Totals
total_revenue <- sum(financial_analysis$Total_Revenue, na.rm = TRUE)
total_gross_margin <- sum(financial_analysis$Gross_Margin_Value, na.rm = TRUE)
total_penalties <- sum(financial_analysis$Penalty_Cost, na.rm = TRUE)

# Retrieve Operational Cost from Task 6 (Optimized Day Shift)
# Ensure you ran Task 6 code so 'total_cost_optimized' exists!
operational_cost_final <- optimized_total_cost

# 4. Calculate Net Profit
net_profit_impact <- total_gross_margin - total_penalties - operational_cost_final

# 5. Build Summary Table
task8_output <- data.frame(
  Category = c(
    "Total Revenue",
    "Gross Margin (Sales Profit)",
    "Less: Late Delivery Penalties",
    "Less: Additional Production Cost",
    "NET PROFIT IMPACT"
  ),
  Amount_INR = c(
    format(total_revenue, big.mark = ","),
    format(total_gross_margin, big.mark = ","),
    paste("-", format(total_penalties, big.mark = ",")),
    paste("-", format(operational_cost_final, big.mark = ",")),
    format(net_profit_impact, big.mark = ",")
  )
)

# Display Task 8
datatable(task8_output, 
          caption = "Task 8: Final Financial Summary",
          options = list(dom = 't', paging = FALSE),
          rownames = FALSE) %>%
  formatStyle(
    'Category',
    target = 'row',
    backgroundColor = styleEqual("NET PROFIT IMPACT", "#ccffcc"),
    fontWeight = styleEqual("NET PROFIT IMPACT", "bold")
  )