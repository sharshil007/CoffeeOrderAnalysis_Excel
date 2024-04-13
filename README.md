# CoffeeOrderAnalysis_Excel
Project: Excel-based Coffee Purchase Analysis

Dataset:

1. Products Sheet:
   - Columns: Product ID (Alphanumeric), Coffee Type (Text), Roast Type (Text), Size (Text), Unit Price (Numeric), Price per 100g (Numeric), Profit (Numeric)
   - Description: Contains information about various coffee products, including their IDs, types, roast types, sizes, prices, and profits.

2. Customers Sheet:
   - Columns: Customer ID (Alphanumeric), Customer Name (Text), Email (Text), Phone Number (Text), Address Line 1 (Text), City (Text), Country (Text), Postcode (Numeric), Loyalty Card (Text: "Yes" or "No")
   - Description: Contains information about customers, including their IDs, names, contact details, addresses, and loyalty card status.

3. Orders Sheet:
   - Columns: Order ID (Alphanumeric), Order Date (Date), Customer ID (Alphanumeric), Product ID (Alphanumeric), Quantity (Numeric)
   - Description: Contains information about orders, including their IDs, dates, customer IDs, product IDs, and quantities.

This dataset provides comprehensive information about coffee products, customers, and orders, making it suitable for analyzing sales, customer behavior, and product performance.

Description: This project involves analyzing a well-curated product dataset containing customer information and their coffee orders. During the project, I enhanced the Orders sheet by adding additional columns derived from the Customers and Products sheets. By employing XLOOKUP and row/column indexing, I extracted the following information:
Additional Columns Added to the Orders Sheet:

Customer Name: The name of the customer corresponding to the Customer ID.
Email: The email address of the customer corresponding to the Customer ID.
Country: The country of the customer corresponding to the Customer ID.
Coffee Type: The type of coffee product corresponding to the Product ID.
Roast Type: The roast type of the coffee product corresponding to the Product ID.
Size: The size of the coffee product corresponding to the Product ID.
Unit Price: The unit price of the coffee product corresponding to the Product ID.
Sales: The total sales amount for each order, calculated based on the quantity and unit price of the product.
This enhancement allowed for a more comprehensive analysis of the dataset, providing additional insights into customer details and product information within the context of each order.

Using advanced Excel functions, I conducted a comprehensive analysis and visualization of customer purchase data.

Key Performance Indicators (KPIs) Focused On:

1. Total Sales Over Time: Utilizing monthly data, I analyzed the total sales trend over time to identify patterns and fluctuations.

2. Country-wise Sales Analysis: I conducted an analysis to determine sales performance across different countries, providing insights into regional trends and preferences.

3. Top 5 Customer Purchase Analysis: I identified and analyzed the top 5 customers based on their purchase behavior, offering valuable insights into customer preferences and loyalty.

4. Enhanced Dashboard Interactivity: I enhanced the dashboard's usability by adding filters for roast type, loyalty card status, and bag size, allowing for easy data filtering and exploration.

Overall, this project showcases Excel's capabilities in data analysis and visualization, demonstrating its effectiveness in deriving insights from complex datasets. The interactive dashboard provides a user-friendly interface for exploring and understanding the data.
