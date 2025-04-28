import streamlit as st
import pandas as pd

# Purpose of the Page
st.title("Excel & Power BI Learning Hub")
st.write("""
    Welcome to the **Excel & Power BI Learning Hub**! 
    This page provides detailed tutorials, tips, and tricks for mastering Excel formulas, 
    Power BI, and improving your data analysis skills. 
    Explore different sections to learn powerful techniques, and get hands-on with interactive demos!
""")

# Add LinkedIn link to Footer
st.markdown("---")
st.write("Connect with me on LinkedIn: [Ashwik Bire](https://linkedin.com/in/ashwik-bire-b2a000186)")

# Step 2: Navigation Sidebar
navigation = st.sidebar.radio("Navigate", [
    "Home",
    "Excel Learning",
    "Power BI Learning",
    "Excel Tips & Tricks",
    "Interactive Demo",
    "Excel Functions",
    "Power BI Dashboard Creation",
    "Advanced Excel Techniques",
    "Power BI Advanced Features",
    "Learning Resources"
])

# Define Content for Different Sections
if navigation == "Home":
    st.header("Welcome to the Learning Hub!")
    st.write("Explore the various tutorials and tips sections, including Excel and Power BI learning, "
             "and enhance your data analysis skills with practical demonstrations.")
    import streamlit as st
    import pandas as pd
    import plotly.express as px

    # Purpose of the Home Page
    st.title("Welcome to the Excel & Power BI Learning Hub!")
    st.write("""
        **Welcome to your one-stop learning hub for mastering Excel and Power BI!**
        This interactive learning space offers comprehensive tutorials, tips, tricks, and hands-on demonstrations to help you build your skills in Excel and Power BI.

        Whether you are just starting or want to dive into advanced techniques, this app has something for everyone. Below is a summary of the key features you will find:
    """)

    # Navigation Instructions
    st.subheader("How to Navigate:")
    st.write("""
        Use the sidebar on the left to explore different sections:
        - **Excel Learning**: Learn about Excel functions, charts, and advanced techniques.
        - **Power BI Learning**: Explore how to create stunning dashboards and reports in Power BI.
        - **Excel Tips & Tricks**: Discover tips to speed up your workflow and improve your Excel efficiency.
        - **Interactive Demos**: Try hands-on demos with Excel formulas like SUM and AVERAGE.
        - **Learning Resources**: Access a curated list of books, YouTube channels, and websites for further learning.

        Select a section from the sidebar, and start learning!
    """)

    # Features Overview with Visuals
    st.subheader("Key Features of the Hub")

    # Visualizing Excel Data Example (Plotting a simple bar chart)
    st.write("Here’s a quick preview of how Excel data can be visualized.")
    df = pd.DataFrame({
        "Categories": ["Sales", "Marketing", "Development", "HR", "Finance"],
        "Amount": [2000, 1500, 3000, 1200, 1800]
    })
    fig = px.bar(df, x="Categories", y="Amount", title="Department Budget Visualization")
    st.plotly_chart(fig)

    # Power BI Visualization Example
    st.write("Check out how powerful Power BI visuals can be.")
    # Simulate a Power BI visual with Plotly (simple bar chart)
    df_powerbi = pd.DataFrame({
        "Region": ["North", "South", "East", "West"],
        "Revenue": [5000, 4000, 4500, 3500]
    })
    fig_powerbi = px.bar(df_powerbi, x="Region", y="Revenue", title="Regional Revenue Performance")
    st.plotly_chart(fig_powerbi)

    # Encouragement to Start Learning
    st.write("""
        Explore the various sections and get started with hands-on learning! 
        The more you practice, the better your data analysis skills will become.
        We have step-by-step guides, tutorials, and demos waiting for you!
    """)

    # LinkedIn Profile Section
    st.markdown("---")
    st.subheader("Connect with Me!")
    st.write("Feel free to connect with me on LinkedIn for more insights, tips, and professional growth:")
    st.write("[Ashwik Bire - LinkedIn Profile](https://linkedin.com/in/ashwik-bire-b2a000186)")

    # Call to Action Button to Explore More
    st.markdown("---")
    st.button("Start Learning Now!", help="Click to explore Excel and Power BI tutorials.")



elif navigation == "Excel Learning":
    st.header("Excel Learning")
    # Add detailed content for Excel Learning here...
    import streamlit as st
    import pandas as pd
    import plotly.express as px

    # Excel Learning Page
    st.title("Excel Learning Hub")

    # Introduction to Excel
    st.header("What is Excel?")
    st.write("""
        **Microsoft Excel** is a powerful spreadsheet application used for data analysis, calculation, visualization, and reporting. 
        It is part of the Microsoft Office Suite and is widely used in businesses, schools, and other professional environments.

        Excel provides:
        - **Data Entry and Formatting**: Enter and format data in rows and columns.
        - **Functions & Formulas**: Perform calculations using built-in functions.
        - **Charts and Graphs**: Visualize your data using various chart types.
        - **PivotTables**: Summarize and analyze large sets of data quickly.
        - **Data Management**: Sort, filter, and manipulate data efficiently.

        Whether you're a beginner or an advanced user, Excel has tools and functions that can help you analyze and present your data effectively.
    """)

    # Excel Basic Functions
    st.header("Basic Excel Functions")
    st.write("""
        Here are some of the most commonly used Excel functions for beginners:

        - **SUM()**: Adds numbers in a range of cells.
        - **AVERAGE()**: Calculates the average of a range of cells.
        - **MIN()** and **MAX()**: Finds the minimum or maximum value in a range.
        - **IF()**: Performs conditional logic based on a specified condition.
        - **COUNT()**: Counts the number of numeric entries in a range.
        - **VLOOKUP()**: Searches for a value in the first column of a range and returns a value in the same row from a specified column.

        Example: 
        ```excel
        =SUM(A1:A10)  # Adds the values in cells A1 through A10
        =AVERAGE(B1:B10)  # Calculates the average of the values in B1 to B10
        ```

        These basic functions are useful for general calculations and data analysis.
    """)

    # Data Visualization in Excel
    st.header("Creating Charts in Excel")
    st.write("""
        Excel provides various types of charts and graphs to visualize your data, including:

        - **Column/Bar Charts**: Good for comparing different categories.
        - **Line Charts**: Great for visualizing trends over time.
        - **Pie Charts**: Best for showing proportions of a whole.
        - **Scatter Plots**: Used to visualize the relationship between two variables.

        To create a chart:
        1. Select the data you want to plot.
        2. Click the **Insert** tab in the ribbon.
        3. Choose a chart type from the **Charts** group.

        Example of a Column Chart:
        """)

    # Create a simple chart using Plotly to show data visualization in Excel (for Streamlit's interactive content)
    df = pd.DataFrame({
        "Category": ["Sales", "Marketing", "Development", "HR", "Finance"],
        "Amount": [2500, 1500, 3000, 1200, 1800]
    })

    fig = px.bar(df, x="Category", y="Amount", title="Department Budget Visualization")
    st.plotly_chart(fig)

    # Data Management in Excel
    st.header("Data Management in Excel")
    st.write("""
        Excel offers powerful data management tools to help you clean, sort, and filter your data:

        - **Sorting**: Organize your data in ascending or descending order.
        - **Filtering**: Display only the data that meets certain criteria.
        - **Find and Replace**: Quickly locate specific values and replace them.
        - **Conditional Formatting**: Highlight data based on specific rules, such as coloring cells based on values.
        - **PivotTables**: Create summaries of large datasets by grouping and aggregating data.

        PivotTables are a great way to analyze and summarize large amounts of data quickly. To create a PivotTable:
        1. Select the data range.
        2. Click on **Insert > PivotTable**.
        3. Choose the fields to analyze and drag them into the PivotTable Field List.
        4. Excel will create a dynamic summary of your data based on your selections.
    """)

    # Advanced Excel Features
    st.header("Advanced Excel Features")
    st.write("""
        For advanced Excel users, here are some features that can take your skills to the next level:

        - **Array Formulas**: Perform multiple calculations on one or more items in an array (e.g., **SUMPRODUCT()**).
        - **Data Validation**: Ensure that data entered into a cell meets certain criteria (e.g., dropdown lists, date ranges).
        - **Named Ranges**: Assign names to ranges to make your formulas easier to read and manage.
        - **Macros**: Automate repetitive tasks by recording actions or writing VBA code.
        - **What-If Analysis**: Use tools like **Goal Seek** or **Data Tables** to analyze different scenarios.

        These advanced features are useful for complex data analysis and automating tasks in Excel.
    """)

    # Excel Shortcuts for Efficiency
    st.header("Excel Shortcuts")
    st.write("""
        Using keyboard shortcuts in Excel can greatly improve your efficiency:

        - **Ctrl + C**: Copy selected cells.
        - **Ctrl + V**: Paste copied cells.
        - **Ctrl + Z**: Undo the last action.
        - **Ctrl + Shift + L**: Add or remove filters.
        - **Ctrl + Arrow keys**: Move to the edge of data ranges.
        - **Alt + E, S, V**: Paste special.
        - **Ctrl + T**: Create a table from a data range.

        Mastering these shortcuts will save you time and make working in Excel faster and more efficient.
    """)

    # Excel Learning Resources
    st.header("Learning Resources for Excel")
    st.write("""
        If you want to learn more about Excel, here are some great resources:

        - **Excel Help & Training (Microsoft)**: [Microsoft Excel Help](https://support.microsoft.com/en-us/excel)
        - **YouTube Channels**:
            - ExcelIsFun: [YouTube Link](https://www.youtube.com/c/ExcelIsFun)
            - Leila Gharani: [YouTube Link](https://www.youtube.com/c/LeilaGharani)
        - **Books**:
            - "Excel 2019 Bible" by Michael Alexander and Richard Kusleika
            - "Excel Formulas and Functions For Dummies" by Ken Bluttman
        - **Online Courses**:
            - Excel Learning Path on [LinkedIn Learning](https://www.linkedin.com/learning/)
            - Excel Courses on [Coursera](https://www.coursera.org/courses?query=excel)

        These resources will help you dive deeper into Excel's features and functions, from basic to advanced levels.
    """)


elif navigation == "Power BI Learning":
    st.header("Power BI Learning")
    #
    import streamlit as st
    import pandas as pd

    # Power BI Learning Page
    st.title("Power BI Learning Hub")

    # Introduction to Power BI
    st.header("What is Power BI?")
    st.write("""
        **Power BI** is a suite of business analytics tools from Microsoft that allows users to visualize their data, share insights, and make data-driven decisions. 
        Power BI is composed of several key components:
        - **Power BI Desktop**: A free application used to create reports and visualizations.
        - **Power BI Service**: A cloud-based platform to publish and share reports and dashboards.
        - **Power BI Mobile**: Mobile apps that allow you to access your reports and dashboards on the go.
        - **Power BI Gateway**: Used for connecting on-premises data sources to Power BI.

        Power BI helps users gather data from multiple sources, perform transformations, and create compelling visual reports.
    """)

    # Power BI Desktop
    st.header("Power BI Desktop: Data Visualization")
    st.write("""
        Power BI Desktop is the primary tool for creating reports and visualizations. It offers a range of features that include:
        - **Data Import**: Connect to multiple data sources (Excel, SQL Server, Web, etc.).
        - **Data Transformation**: Clean and shape your data using Power Query Editor.
        - **Data Modeling**: Build relationships between different data tables to create models.
        - **Creating Reports**: Use visuals like charts, tables, and maps to represent data.

        **Steps to create a simple report in Power BI Desktop**:
        1. Open Power BI Desktop and click on **Get Data** to connect to your data source.
        2. Use the **Query Editor** to clean and transform your data.
        3. Build a model by creating relationships between data tables.
        4. Add visuals by dragging fields onto the report canvas.
        5. Customize the appearance of the visuals and add interactivity.

        Once the report is ready, you can publish it to the **Power BI Service** for sharing and collaboration.
    """)

    # Power Query Tutorial
    st.header("Power Query: Data Transformation")
    st.write("""
        Power Query is a tool within Power BI that allows you to transform your data before loading it into your reports. Common tasks include:
        - Removing duplicates
        - Filtering data
        - Changing column types
        - Merging multiple data sources
        - Creating calculated columns

        **Example of Power Query Transformation**:
        1. Import data into Power Query.
        2. Clean the data by removing rows with null values or applying filters.
        3. Combine multiple tables using the **Merge Queries** function.
        4. Load the transformed data into Power BI Desktop.

        Power Query also supports **M code** for advanced transformations.
    """)

    # DAX Formulas Introduction
    st.header("DAX Formulas: Advanced Data Analysis")
    st.write("""
        **DAX (Data Analysis Expressions)** is a formula language used to perform calculations in Power BI. It allows users to create custom calculations for measures, columns, and tables.

        **Basic DAX Functions**:
        - **SUM()**: Adds up values in a column.
        - **AVERAGE()**: Calculates the average of a column.
        - **COUNTROWS()**: Counts the number of rows in a table.
        - **IF()**: Creates conditional expressions.
        - **CALCULATE()**: Modifies the context in which a calculation is performed.

        **Example**: Create a measure to calculate total sales:
        ```DAX
        Total Sales = SUM(Sales[Amount])
        ```

        More advanced formulas allow users to work with time intelligence, such as calculating **Year-to-Date** (YTD) values and **Month-over-Month** (MoM) changes.
    """)

    # Power BI Service: Sharing and Collaboration
    st.header("Power BI Service: Publishing & Sharing Reports")
    st.write("""
        After creating reports in Power BI Desktop, you can publish them to the **Power BI Service** to share and collaborate with others:

        **Steps to Publish a Report**:
        1. Click **Publish** in Power BI Desktop.
        2. Sign in to your Power BI account (or create one if you don’t have it).
        3. Choose the workspace where you want to publish the report.

        Once the report is published, you can:
        - Share reports with colleagues.
        - Collaborate on dashboards.
        - Set up automatic data refresh.

        The **Power BI Service** is cloud-based, so you can access your reports from anywhere, including mobile devices.

        Power BI Service also allows users to set up **Row-Level Security (RLS)** to restrict data access based on user roles.
    """)

    # Power BI Mobile App
    st.header("Power BI Mobile App")
    st.write("""
        The **Power BI Mobile App** allows you to access and interact with your reports and dashboards on the go. You can:
        - View reports and dashboards.
        - Receive data alerts on your mobile device.
        - Share insights with your team directly from the app.

        This mobile access is great for quick decisions and staying updated with the latest data while on the move.
    """)

    # Learning Resources for Power BI
    st.header("Learning Resources for Power BI")
    st.write("""
        Here are some great resources to help you learn more about Power BI:
        - **Power BI Documentation**: [Microsoft Power BI Docs](https://docs.microsoft.com/en-us/power-bi/)
        - **YouTube Channels**:
          - Guy in a Cube: [YouTube Link](https://www.youtube.com/c/GuyInACube)
          - Enterprise DNA: [YouTube Link](https://www.youtube.com/c/EnterpriseDNA)
        - **Books**:
          - "The Definitive Guide to DAX" by Marco Russo & Alberto Ferrari
          - "Power BI Step-by-Step" by Mike Alexander and Jared Decker
        - **Online Courses**:
          - Power BI Learning Path on [LinkedIn Learning](https://www.linkedin.com/learning/)
          - Power BI Courses on [Coursera](https://www.coursera.org/)

        These resources will take your Power BI skills to the next level and keep you up to date with the latest features.
    """)


elif navigation == "Excel Tips & Tricks":
    st.header("Excel Tips & Tricks")
    # Add detailed content for Excel Tips & Tricks here...
    import streamlit as st

    # Excel Tips & Tricks Page
    st.title("Excel Tips & Tricks")

    # Keyboard Shortcuts
    st.header("Keyboard Shortcuts")
    st.write("""
        Excel keyboard shortcuts can greatly enhance your productivity. Here are some of the most useful ones:

        - **Ctrl + C**: Copy selected cells.
        - **Ctrl + X**: Cut selected cells.
        - **Ctrl + V**: Paste copied/cut cells.
        - **Ctrl + Z**: Undo the last action.
        - **Ctrl + Y**: Redo the last undone action.
        - **Ctrl + Arrow keys**: Move to the edge of the data range.
        - **Ctrl + Shift + L**: Toggle filters on and off.
        - **Ctrl + T**: Create a table from selected data.
        - **Ctrl + 1**: Open the Format Cells dialog box (for formatting options).
        - **Alt + E, S, V**: Paste Special options.
        - **F2**: Edit the selected cell directly in the formula bar.

        These shortcuts can save you a lot of time when navigating and working in Excel.
    """)

    # Cell Formatting Tricks
    st.header("Cell Formatting Tricks")
    st.write("""
        Excel provides a range of formatting options to make your data more readable and professional:

        - **Quick Formatting**: Select a range and press **Ctrl + Shift + $** to apply currency format.
        - **Merge & Center**: Use **Ctrl + Alt + M** to merge cells and center the text.
        - **Conditional Formatting**: Highlight data based on rules (e.g., highlight cells greater than a certain value):
            1. Select the range.
            2. Go to **Home > Conditional Formatting**.
            3. Choose your formatting rule.
        - **Format as Table**: Use **Ctrl + T** to quickly convert a range of data into a table with filter options.
        - **Format Painter**: Use the **Format Painter** (the paintbrush icon) to copy the formatting of one cell to another.

        These tricks can help make your data visually appealing and easier to read.
    """)

    # Hidden Features in Excel
    st.header("Hidden Features")
    st.write("""
        Excel has several hidden features that are not immediately obvious but can be incredibly useful:

        - **Flash Fill**: Automatically fill in values based on patterns. For example, if you start typing an email address or name, Excel will predict and fill the remaining values:
            1. Start typing a value.
            2. Press **Ctrl + E** to use Flash Fill.
        - **Go To Special**: Quickly select specific types of cells (e.g., blank cells, formulas, etc.):
            1. Press **F5** or **Ctrl + G**.
            2. Click on **Special...** to choose the type of cells to select.
        - **Freeze Panes**: Freeze rows or columns so they stay visible when scrolling:
            1. Select the row/column you want to freeze.
            2. Go to **View > Freeze Panes**.
        - **Group and Ungroup**: Organize data into collapsible groups:
            1. Select the rows/columns.
            2. Go to **Data > Group**.
        - **Quick Analysis**: Quickly analyze your data with Excel’s built-in tools:
            1. Select the range of data.
            2. Click the **Quick Analysis** button that appears at the bottom right.

        These hidden features are perfect for streamlining your work and making tasks easier.
    """)

    # Lesser-Known Formulas and Functions
    st.header("Formulas and Functions")
    st.write("""
        These lesser-known formulas can help you perform advanced calculations in Excel:

        - **INDEX() and MATCH()**: A powerful alternative to **VLOOKUP()** for looking up data.
            Example: 
            ```excel
            =INDEX(A2:A10, MATCH("Sales", B2:B10, 0))
            ```
        - **TEXT()**: Convert values to text with custom formatting. 
            Example: 
            ```excel
            =TEXT(A1, "dd/mm/yyyy")
            ```
        - **SUMIF() / COUNTIF()**: Sum or count based on specific criteria.
            Example:
            ```excel
            =SUMIF(A1:A10, ">1000")
            ```
        - **IFERROR()**: Handle errors in formulas.
            Example:
            ```excel
            =IFERROR(A1/B1, "Error")
            ```
        - **TRANSPOSE()**: Change the orientation of data from rows to columns or vice versa.
            Example:
            ```excel
            =TRANSPOSE(A1:B3)
            ```

        These formulas can save you time and help you tackle complex data analysis tasks.
    """)

    # Data Management Tips
    st.header("Data Management")
    st.write("""
        Excel offers several tips for managing large datasets:

        - **Data Validation**: Restrict data entry to specific criteria (e.g., dates, numbers).
            1. Select the range.
            2. Go to **Data > Data Validation**.
        - **Remove Duplicates**: Eliminate duplicate values in a range.
            1. Select the range.
            2. Go to **Data > Remove Duplicates**.
        - **Text to Columns**: Split data in a cell into multiple columns.
            1. Select the range.
            2. Go to **Data > Text to Columns**.
        - **Power Query**: Use Power Query to import, clean, and transform data from various sources.
        - **Filters**: Quickly filter data by selecting the drop-down arrows in column headers (requires **Ctrl + Shift + L**).

        These tips can help you stay organized when working with large datasets.
    """)

    # Automation Tips
    st.header("Automation in Excel")
    st.write("""
        Automating repetitive tasks in Excel can save you a lot of time:

        - **Recording Macros**: Automate a series of actions by recording a macro:
            1. Go to **View > Macros > Record Macro**.
            2. Perform the tasks you want to automate.
            3. Stop recording the macro.
            4. Assign the macro to a button or keyboard shortcut.
        - **VBA (Visual Basic for Applications)**: Write custom VBA scripts to automate complex workflows.
        - **Dynamic Arrays**: Use dynamic array functions like **FILTER()**, **SORT()**, and **UNIQUE()** to automate and simplify your data management.

        Automation can reduce manual work and improve the consistency of your tasks.
    """)

    # Customizing Excel
    st.header("Customizing Excel")
    st.write("""
        Customize Excel to better suit your needs:

        - **Quick Access Toolbar**: Add your most frequently used commands for easy access.
            1. Right-click on any command and select **Add to Quick Access Toolbar**.
        - **Ribbon Customization**: Modify the Ribbon to add or remove tabs and groups.
            1. Go to **File > Options > Customize Ribbon**.
        - **Themes and Colors**: Change Excel's appearance by adjusting themes, fonts, and colors.
            1. Go to **File > Options > General > Personalize your copy of Microsoft Office**.

        Customizing Excel will make it more tailored to your workflow.
    """)


elif navigation == "Interactive Demo":
    st.header("Interactive Demo: Excel Formulas in Action")

    st.subheader("SUM Function Demo")
    st.write("Enter a set of numbers, and we'll calculate the sum for you!")

    # User input for numbers
    numbers = st.text_area("Enter numbers (comma separated):", "1, 2, 3, 4, 5")
    numbers = [int(i.strip()) for i in numbers.split(",") if i.strip().isdigit()]

    if numbers:
        total_sum = sum(numbers)
        st.write(f"Sum of numbers: {total_sum}")

    st.markdown("---")

    st.subheader("AVERAGE Function Demo")
    st.write("Enter a set of numbers, and we'll calculate the average for you!")

    # User input for numbers for average calculation
    numbers_avg = st.text_area("Enter numbers (comma separated):", "10, 20, 30, 40, 50")
    numbers_avg = [int(i.strip()) for i in numbers_avg.split(",") if i.strip().isdigit()]

    if numbers_avg:
        average = sum(numbers_avg) / len(numbers_avg)
        st.write(f"Average of numbers: {average}")

    st.markdown("---")

elif navigation == "Excel Functions":
    import streamlit as st

    # Excel Functions Page
    st.title("Excel Functions Guide")

    # Arithmetic Functions
    st.header("Arithmetic Functions")
    st.write("""
        Excel provides several basic arithmetic functions that help with simple calculations:

        - **SUM()**: Adds numbers in a range of cells.
            Example: 
            ```excel
            =SUM(A1:A10)  # Adds the values in cells A1 through A10
            ```
        - **SUBTOTAL()**: Calculates subtotals for data in a table or range.
            Example: 
            ```excel
            =SUBTOTAL(9, A1:A10)  # The "9" argument calculates the sum
            ```
        - **PRODUCT()**: Multiplies all the numbers in a range.
            Example: 
            ```excel
            =PRODUCT(A1:A5)  # Multiplies the values in cells A1 to A5
            ```
        - **AVERAGE()**: Calculates the average of a range of numbers.
            Example:
            ```excel
            =AVERAGE(A1:A10)  # Returns the average of cells A1 through A10
            ```
        - **MIN() / MAX()**: Returns the minimum or maximum value in a range.
            Example:
            ```excel
            =MIN(A1:A10)  # Returns the smallest number in the range
            =MAX(A1:A10)  # Returns the largest number in the range
            ```

        These arithmetic functions are the building blocks of most calculations in Excel.
    """)

    # Lookup & Reference Functions
    st.header("Lookup & Reference Functions")
    st.write("""
        These functions help you search for values in a range or table and return related information:

        - **VLOOKUP()**: Searches for a value in the first column of a table and returns a value in the same row from another column.
            Example:
            ```excel
            =VLOOKUP(A1, B2:D10, 3, FALSE)  # Looks for value in A1, searches in columns B to D, returns value from 3rd column
            ```
        - **HLOOKUP()**: Similar to **VLOOKUP**, but it searches for a value in the first row of a table and returns a value in the same column from another row.
            Example:
            ```excel
            =HLOOKUP(A1, A2:F6, 3, FALSE)  # Looks for value in A1, searches in rows A to F, returns value from 3rd row
            ```
        - **INDEX()**: Returns the value of a cell in a specific row and column within a range.
            Example:
            ```excel
            =INDEX(A1:B10, 2, 1)  # Returns value in 2nd row, 1st column
            ```
        - **MATCH()**: Returns the position of a value in a range.
            Example:
            ```excel
            =MATCH("Apple", A1:A10, 0)  # Finds the position of "Apple" in the range A1 to A10
            ```
        - **OFFSET()**: Returns a value based on a reference cell and offset from that reference.
            Example:
            ```excel
            =OFFSET(A1, 2, 3)  # Returns the value 2 rows down and 3 columns to the right of cell A1
            ```

        These functions are useful when you need to find and return specific data from a table or range.
    """)

    # Logical Functions
    st.header("Logical Functions")
    st.write("""
        Logical functions help you make decisions based on conditions:

        - **IF()**: Performs a conditional test and returns one value if true and another if false.
            Example:
            ```excel
            =IF(A1 > 10, "High", "Low")  # Returns "High" if value in A1 is greater than 10, else "Low"
            ```
        - **AND() / OR()**: Returns TRUE if all conditions are true (**AND**) or if at least one condition is true (**OR**).
            Example:
            ```excel
            =AND(A1 > 10, B1 < 5)  # Returns TRUE if both conditions are true
            =OR(A1 > 10, B1 < 5)  # Returns TRUE if at least one condition is true
            ```
        - **IFERROR()**: Returns a custom value if a formula returns an error, otherwise returns the formula result.
            Example:
            ```excel
            =IFERROR(A1/B1, "Error")  # If division results in an error, returns "Error"
            ```
        - **NOT()**: Reverses the logic of a condition.
            Example:
            ```excel
            =NOT(A1 > 10)  # Returns TRUE if A1 is not greater than 10
            ```

        These functions are essential for performing conditional logic in Excel formulas.
    """)

    # Text Functions
    st.header("Text Functions")
    st.write("""
        Excel has many text functions for manipulating text strings:

        - **CONCATENATE()**: Joins two or more strings into one.
            Example:
            ```excel
            =CONCATENATE(A1, " ", B1)  # Joins values from A1 and B1 with a space between them
            ```
        - **TEXT()**: Formats a number as text with a specific format.
            Example:
            ```excel
            =TEXT(A1, "dd/mm/yyyy")  # Formats the date in A1 as day/month/year
            ```
        - **LEFT() / RIGHT()**: Extracts a specific number of characters from the left or right of a string.
            Example:
            ```excel
            =LEFT(A1, 3)  # Extracts the first 3 characters from A1
            =RIGHT(A1, 3)  # Extracts the last 3 characters from A1
            ```
        - **MID()**: Extracts a specific number of characters from the middle of a string.
            Example:
            ```excel
            =MID(A1, 2, 3)  # Extracts 3 characters from A1 starting from the 2nd character
            ```
        - **TRIM()**: Removes leading and trailing spaces from a text string.
            Example:
            ```excel
            =TRIM(A1)  # Removes spaces before and after the text in A1
            ```

        These text functions help you clean and manipulate text data.
    """)

    # Date & Time Functions
    st.header("Date & Time Functions")
    st.write("""
        Excel offers various functions for working with dates and times:

        - **TODAY()**: Returns the current date.
            Example:
            ```excel
            =TODAY()  # Returns the current date
            ```
        - **NOW()**: Returns the current date and time.
            Example:
            ```excel
            =NOW()  # Returns the current date and time
            ```
        - **DATEDIF()**: Calculates the difference between two dates.
            Example:
            ```excel
            =DATEDIF(A1, B1, "Y")  # Returns the number of years between the dates in A1 and B1
            ```
        - **YEAR() / MONTH() / DAY()**: Extracts the year, month, or day from a date.
            Example:
            ```excel
            =YEAR(A1)  # Extracts the year from the date in A1
            =MONTH(A1)  # Extracts the month from the date in A1
            =DAY(A1)  # Extracts the day from the date in A1
            ```
        - **NETWORKDAYS()**: Returns the number of working days between two dates.
            Example:
            ```excel
            =NETWORKDAYS(A1, B1)  # Returns the number of working days between the dates in A1 and B1
            ```

        These date and time functions are helpful when calculating durations or extracting specific components of a date.
    """)

    # Array Functions
    st.header("Array Functions")
    st.write("""
        Array functions allow you to perform operations on an entire array of data:

        - **TRANSPOSE()**: Changes the orientation of a range from rows to columns or vice versa.
            Example:
            ```excel
            =TRANSPOSE(A1:B10)  # Changes the data from a vertical to horizontal orientation
            ```
        - **FREQUENCY()**: Calculates how often values occur in a range and returns a vertical array of numbers.
            Example:
            ```excel
            =FREQUENCY(A1:A10, B1:B5)  # Returns an array of frequencies for the data in A1:A10
            ```

        Array functions are useful when working with ranges and arrays of data.
    """)

    # Financial Functions
    st.header("Financial Functions")
    st.write("""
        Excel provides several financial functions that are useful for investment and financial analysis:

        - **PMT()**: Calculates the payment for a loan based on constant payments and a constant interest rate.
            Example:
            ```excel
            =PMT(5%/12, 60, -10000)  # Calculates monthly payment for a 5% annual rate over 60 months for a $10,000 loan
            ```
        - **FV()**: Calculates the future value of an investment based on periodic, constant payments and a constant interest rate.
            Example:
            ```excel
            =FV(5%/12, 60, -200, 0)  # Calculates the future value of a $200 monthly investment at 5% annual interest over 5 years
            ```

        These functions are helpful for analyzing loans, investments, and other financial scenarios.
    """)

    # Statistical Functions
    st.header("Statistical Functions")
    st.write("""
        Excel also provides a variety of statistical functions for analyzing data:

        - **AVERAGEIF()**: Calculates the average of cells that meet a specified condition.
            Example:
            ```excel
            =AVERAGEIF(A1:A10, ">10")  # Returns the average of values greater than 10
            ```
        - **STDEV()**: Calculates the standard deviation of a range of numbers.
            Example:
            ```excel
            =STDEV(A1:A10)  # Returns the standard deviation of values in A1 to A10
            ```
        - **COUNTIF()**: Counts the number of cells that meet a specified condition.
            Example:
            ```excel
            =COUNTIF(A1:A10, ">10")  # Counts the number of cells greater than 10
            ```

        Statistical functions help you analyze and summarize data in Excel.
    """)



elif navigation == "Advanced Excel Techniques":
    st.header("Advanced Excel Techniques")
    import streamlit as st

    # Power BI Dashboard creation page
    st.title("Power BI Dashboard Creation Guide")

    # Introduction Section
    st.write("""
        Power BI dashboards allow you to present data in a visually interactive way. With Power BI, you can:
        - Create bar charts, pie charts, line graphs, and more.
        - Filter and drill down on specific aspects of your data.
        - Create KPIs to highlight key metrics.
        - Share dashboards for collaboration and real-time reporting.
    """)

    # Steps Overview
    st.header("Steps to Create a Power BI Dashboard")
    st.write("""
    1. **Prepare Your Data**: Import data from Excel, SQL, or other sources.
    2. **Create Relationships**: Define relationships between tables.
    3. **Create Visualizations**: Use various chart types such as bar, line, and pie.
    4. **Add Interactivity**: Use slicers and filters to allow users to interact with the data.
    5. **Design Layout**: Arrange visuals to make the dashboard user-friendly.
    6. **Publish & Share**: Publish your dashboard to Power BI Service and share with others.
    """)

    # Example Scenario Section
    st.header("Example: Sales Dashboard")
    st.write("""
        Imagine you have sales data with fields such as `Product`, `Sales Amount`, `Region`, and `Sales Date`. Here’s how you can use Power BI to build a dashboard:

        - **Total Sales**: A **KPI Card** to show the total sales amount.
        - **Sales by Region**: A **Bar Chart** showing total sales by region.
        - **Sales Trend**: A **Line Chart** showing sales over time.
        - **Product Breakdown**: A **Pie Chart** showing sales by product category.

        Add slicers to filter the data by date or region, making the dashboard interactive.
    """)

    # Conclusion Section
    st.header("Conclusion")
    st.write("""
        Power BI is a powerful tool for creating data dashboards that provide actionable insights. Whether you're analyzing sales data, financial metrics, or any other business data, Power BI offers an intuitive interface to create dynamic and insightful dashboards.
    """)


elif navigation == "Power BI Advanced Features":
    st.header("Power BI Advanced Features")
    import streamlit as st
    import pandas as pd
    import numpy as np
    import matplotlib.pyplot as plt
    import seaborn as sns

    # Sample data
    data = {
        'Region': ['North', 'South', 'East', 'West'],
        'Sales': [25000, 22000, 31000, 15000],
        'Profit': [8000, 5000, 10000, 4000]
    }
    df = pd.DataFrame(data)

    # Title
    st.title("Advanced Power BI Dashboard - Streamlit")

    # Sidebar for user interaction
    region_filter = st.sidebar.selectbox("Select Region", df['Region'].unique())
    filtered_data = df[df['Region'] == region_filter]

    # Visualizations
    st.subheader(f"Sales and Profit for {region_filter} Region")

    # Bar chart for Sales vs Profit
    fig, ax = plt.subplots()
    ax.bar(filtered_data['Region'], filtered_data['Sales'], label="Sales", color="b", alpha=0.6)
    ax.bar(filtered_data['Region'], filtered_data['Profit'], label="Profit", color="g", alpha=0.6)
    ax.set_ylabel('Amount')
    ax.set_title(f'Sales vs Profit - {region_filter}')
    ax.legend()
    st.pyplot(fig)

    # Dynamic Title
    st.subheader(f"Dynamic Data for {region_filter} Region")
    st.write(f"Sales: {filtered_data['Sales'].values[0]}, Profit: {filtered_data['Profit'].values[0]}")

    # What-if simulation (Simple)
    discount = st.slider("Select Discount Rate", 0, 20, 10)
    simulated_sales = filtered_data['Sales'].values[0] * (1 - discount / 100)
    st.write(f"Simulated Sales with {discount}% Discount: {simulated_sales:.2f}")


elif navigation == "Learning Resources":
    st.header("Learning Resources")
    st.write("""
        Here are some valuable learning resources to further enhance your Excel and Power BI skills:
        - **YouTube Channels**:
          - ExcelIsFun: [YouTube Link](https://www.youtube.com/user/ExcelIsFun)
          - Guy in a Cube: [YouTube Link](https://www.youtube.com/c/GuyInACube)
        - **Websites**:
          - ExcelJet: [ExcelJet](https://exceljet.net/)
          - PowerBI Community: [Power BI Community](https://community.powerbi.com/)
        - **Books**:
          - Excel Bible by John Walkenbach
          - The Definitive Guide to DAX by Marco Russo and Alberto Ferrari
    """)
    st.markdown("---")
