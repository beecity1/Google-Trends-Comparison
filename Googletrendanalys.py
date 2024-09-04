import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pytrends.request import TrendReq
import pandas as pd
import plotly.graph_objs as go
import plotly.io as pio
import plotly.express as px
import webbrowser
import time
import math
import os
from bs4 import BeautifulSoup
import json
import re

# Pagination parameters for the HTML output
items_per_page = 100

# Function to calculate averages
def calculate_averages(data, term):
    if data[term].isnull().all() or data[term].sum() == 0:
        print(f"Warning: No data found for term '{term}' globally.")
        return 0, 0, 0, 0  # Return zeros if no data

    # Proceed with normal calculation if data is available
    current_year = data.index[-1].year
    current_month = data.index[-1].month
    current_week = data.index[-1].isocalendar()[1]

    year_data = data[data.index.year == current_year]
    month_data = data[(data.index.year == current_year) & (data.index.month == current_month)]
    week_data = data[(data.index.year == current_year) & (data.index.isocalendar().week == current_week)]

    last_year_data = data[data.index.year == current_year - 1]
    last_year_avg = last_year_data[term].mean() if not last_year_data.empty else year_data[term].mean()

    year_avg = year_data[term].mean()
    month_avg = month_data[term].mean()
    week_avg = week_data[term].mean()

    return year_avg, month_avg, week_avg, last_year_avg


# Function to fetch and plot Google Trends data
# Function to fetch and plot Google Trends data
def plot_trends(search_terms, region1=None, region2=None, generate_map=False):
    global stats

    try:
        if len(search_terms) < 1:
            messagebox.showerror("Error", "Please enter or upload at least one search term.")
            return

        # Remove duplicates
        search_terms = list(dict.fromkeys(search_terms))

        pytrends = TrendReq(hl='en-US', tz=360)
        all_data = pd.DataFrame()
        retry_time = 60  # Initial retry time in seconds
        max_retries = 5  # Maximum number of retries
        missing_terms = []

        # Fetch data for the first region (e.g., US)
        region1_data = pd.DataFrame()
        if region1:
            for i in range(0, len(search_terms), 5):
                group = search_terms[i:i+5]
                for attempt in range(max_retries):
                    try:
                        pytrends.build_payload(group, cat=0, timeframe='today 12-m', geo=region1, gprop='')
                        data = pytrends.interest_over_time()
                        region1_data = pd.concat([region1_data, data[group]], axis=1)
                        break  # Exit loop when successful
                    except Exception as e:
                        if "429" in str(e):
                            print(f"Rate limit exceeded. Waiting for {retry_time} seconds before retrying...")
                            time.sleep(retry_time)
                            retry_time *= 2  # Exponentially increase the wait time
                        else:
                            print(f"Error with terms {group}: {str(e)}. Skipping these terms.")
                            for term in group:
                                missing_terms.append(term)
                            break  # Skip the group and continue to the next one

                time.sleep(2)  # Introduce a delay between requests to avoid hitting the rate limit

        # Fetch data for the second region (e.g., GB)
        region2_data = pd.DataFrame()
        if region2:
            for i in range(0, len(search_terms), 5):
                group = search_terms[i:i+5]
                for attempt in range(max_retries):
                    try:
                        pytrends.build_payload(group, cat=0, timeframe='today 12-m', geo=region2, gprop='')
                        data = pytrends.interest_over_time()
                        region2_data = pd.concat([region2_data, data[group]], axis=1)
                        break  # Exit loop when successful
                    except Exception as e:
                        if "429" in str(e):
                            print(f"Rate limit exceeded. Waiting for {retry_time} seconds before retrying...")
                            time.sleep(retry_time)
                            retry_time *= 2  # Exponentially increase the wait time
                        else:
                            print(f"Error with terms {group}: {str(e)}. Skipping these terms.")
                            for term in group:
                                missing_terms.append(term)
                            break  # Skip the group and continue to the next one

                time.sleep(2)  # Introduce a delay between requests to avoid hitting the rate limit

        # Fetch global data for comparison
        for i in range(0, len(search_terms), 5):
            group = search_terms[i:i+5]
            for attempt in range(max_retries):
                try:
                    pytrends.build_payload(group, cat=0, timeframe='today 12-m', geo='', gprop='')
                    data = pytrends.interest_over_time()
                    all_data = pd.concat([all_data, data[group]], axis=1)
                    break  # Exit loop when successful
                except Exception as e:
                    if "429" in str(e):
                        print(f"Rate limit exceeded. Waiting for {retry_time} seconds before retrying...")
                        time.sleep(retry_time)
                        retry_time *= 2  # Exponentially increase the wait time
                    else:
                        print(f"Error with terms {group}: {str(e)}. Skipping these terms.")
                        for term in group:
                            missing_terms.append(term)
                        break  # Skip the group and continue to the next one

            time.sleep(2)  # Introduce a delay between requests to avoid hitting the rate limit

        all_data = all_data.loc[:, ~all_data.columns.duplicated()]

        stats = []  # Initialize the stats variable

        for term in search_terms:
            if term in all_data.columns:
                try:
                    year_avg, month_avg, week_avg, last_year_avg = calculate_averages(all_data, term)
                    latest_value = all_data[term].iloc[-1]

                    # Fetch averages for Region 1
                    if region1 and term in region1_data.columns:
                        region1_year_avg, region1_month_avg, region1_week_avg = get_selected_region_averages(region1_data, term)
                    else:
                        region1_year_avg = region1_month_avg = region1_week_avg = None

                    # Fetch averages for Region 2
                    if region2 and term in region2_data.columns:
                        region2_year_avg, region2_month_avg, region2_week_avg = get_selected_region_averages(region2_data, term)
                    else:
                        region2_year_avg = region2_month_avg = region2_week_avg = None

                    stats.append((term, latest_value, year_avg, month_avg, week_avg, last_year_avg, region1_year_avg, region1_month_avg, region1_week_avg, region2_year_avg, region2_month_avg, region2_week_avg))
                except Exception as e:
                    print(f"Error processing term '{term}': {str(e)}. Skipping this term.")
                    missing_terms.append(term)

        stats.sort(key=lambda x: x[1], reverse=True)

        # Generate HTML output
        generate_html_output(all_data, search_terms, missing_terms, region1, region1_data, region2, region2_data, generate_map)
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Function to generate HTML output
def generate_html_output(data, search_terms, missing_terms, region1, region1_data, region2, region2_data, generate_map):
    try:
        # Generate the main graph
        fig = go.Figure()

        for term in search_terms:
            if term in data.columns:
                fig.add_trace(go.Scatter(x=data.index, y=data[term], mode='lines+markers', name=term,
                                         hoverinfo='name+y',
                                         marker=dict(size=8),
                                         line=dict(width=2)))

        fig.update_layout(
            title='Google Trends Comparison',
            xaxis_title='Date',
            yaxis_title='Search Interest',
            hovermode='x unified',
            margin=dict(l=40, r=40, t=40, b=40),
            autosize=True,
            showlegend=True,
            hoverlabel=dict(
                bgcolor='white',
                bordercolor='black',
                font_size=12,
                namelength=0
            ),
            dragmode=False,
        )

        fig.update_traces(
            hoverlabel=dict(
                bgcolor="white",
                font_size=12,
                font_family="Arial",
                bordercolor="black"
            ),
            hovertemplate="<b>%{y}</b><extra>%{fullData.name}</extra>"
        )

        graph_html = pio.to_html(fig, full_html=False)

        missing_terms_html = ""
        if missing_terms:
            missing_terms_html = f"""
            <div style="border: 2px solid red; color: red; padding: 10px; margin-bottom: 20px;">
                <strong>These terms encountered issues or were not found in the data:</strong> {', '.join(missing_terms)}
            </div>
            """

        total_pages = math.ceil(len(stats) / items_per_page)
        pagination_html = f"""
        <div style="text-align: center; margin-top: 20px;">
            <button onclick="prevPage()" style="padding: 10px; font-size: 16px;">Previous</button>
            <span id="pageInfo">Page 1 of {total_pages}</span>
            <button onclick="nextPage()" style="padding: 10px; font-size: 16px;">Next</button>
        </div>
        """

        # Header with updated column layout
        stats_html = f"""
        <table id="statsTable" style="width: 100%; border-collapse: collapse;">
        <thead>
            <tr>
                <th onclick="sortTable(0, 'str')">Term <span></span></th>
                <th onclick="sortTable(1, 'num')">Latest Value <span></span></th>
                <th onclick="sortTable(2, 'num')">Year Avg <span></span></th>
                <th onclick="sortTable(3, 'num')">Month Avg <span></span></th>
                <th onclick="sortTable(4, 'num')">Week Avg <span></span></th>
                <th onclick="sortTable(5, 'num')">{region1} Year Avg <span></span></th>
                <th onclick="sortTable(6, 'num')">{region1} Month Avg <span></span></th>
                <th onclick="sortTable(7, 'num')">{region1} Week Avg <span></span></th>
                <th onclick="sortTable(8, 'str')">Top Region <span></span></th>
                <th onclick="sortTable(9, 'num')">Top Region Year Avg <span></span></th>
                <th onclick="sortTable(10, 'num')">Top Region Month Avg <span></span></th>
                <th onclick="sortTable(11, 'num')">Top Region Week Avg <span></span></th>
            </tr>
        </thead>
        <tbody id="statsTableBody">
        """

        for i, stat in enumerate(stats):
            latest_color, latest_arrow = compare_values(stat[1], stat[2])
            year_color, year_arrow = compare_values(stat[2], stat[5])
            month_color, month_arrow = compare_values(stat[3], stat[2])
            week_color, week_arrow = compare_values(stat[4], stat[3])

            stats_html += f"""
            <tr data-index="{i + 1}">
                <td style="padding: 8px;">{stat[0]}</td>
                <td style="padding: 8px; color: {latest_color};">{stat[1]:.2f} {latest_arrow}</td>
                <td style="padding: 8px; color: {year_color};">{stat[2]:.2f} {year_arrow}</td>
                <td style="padding: 8px; color: {month_color};">{stat[3]:.2f} {month_arrow}</td>
                <td style="padding: 8px; color: {week_color};">{stat[4]:.2f} {week_arrow}</td>
            """

            if region1_data is not None and not region1_data.empty:
                selected_region1_year_avg, selected_region1_month_avg, selected_region1_week_avg = get_selected_region_averages(region1_data, stat[0])
                selected1_year_color, selected1_year_arrow = compare_values(selected_region1_year_avg, stat[2])
                selected1_month_color, selected1_month_arrow = compare_values(selected_region1_month_avg, stat[3])
                selected1_week_color, selected1_week_arrow = compare_values(selected_region1_week_avg, stat[4])

                stats_html += f"""
                    <td style="padding: 8px; color: {selected1_year_color};">{selected_region1_year_avg:.2f} {selected1_year_arrow}</td>
                    <td style="padding: 8px; color: {selected1_month_color};">{selected_region1_month_avg:.2f} {selected1_month_arrow}</td>
                    <td style="padding: 8px; color: {selected1_week_color};">{selected_region1_week_avg:.2f} {selected1_week_arrow}</td>
                """

            top_region_name, top_region_year_avg, top_region_month_avg, top_region_week_avg = get_top_region_comparison_data(stat[0])
            top_region_year_color, top_region_year_arrow = compare_values(top_region_year_avg, stat[2])
            top_region_month_color, top_region_month_arrow = compare_values(top_region_month_avg, stat[3])
            top_region_week_color, top_region_week_arrow = compare_values(top_region_week_avg, stat[4])

            # Ensure proper encoding of region names
            top_region_name = top_region_name.encode('latin1').decode('latin1')

            stats_html += f"""
                <td style="padding: 8px;">{top_region_name}</td>
                <td style="padding: 8px; color: {top_region_year_color};">{top_region_year_avg:.2f} {top_region_year_arrow}</td>
                <td style="padding: 8px; color: {top_region_month_color};">{top_region_month_avg:.2f} {top_region_month_arrow}</td>
                <td style="padding: 8px; color: {top_region_week_color};">{top_region_week_avg:.2f} {top_region_week_arrow}</td>
            </tr>
            """

        stats_html += """
        </tbody>
        </table>
        <script>
        let currentPage = 1;
        const itemsPerPage = """ + str(items_per_page) + """;
        const totalItems = """ + str(len(stats)) + """;
        const totalPages = """ + str(total_pages) + """;

        function renderTablePage(page) {
            const rows = document.querySelectorAll('#statsTableBody tr');
            rows.forEach(row => {
                const rowIndex = parseInt(row.getAttribute('data-index'));
                if (rowIndex > (page - 1) * itemsPerPage && rowIndex <= page * itemsPerPage) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
            document.getElementById('pageInfo').textContent = `Page ${page} of ${totalPages}`;
        }

        function nextPage() {
            if (currentPage < totalPages) {
                currentPage++;
                renderTablePage(currentPage);
            }
        }

        function prevPage() {
            if (currentPage > 1) {
                currentPage--;
                renderTablePage(currentPage);
            }
        }

        function sortTable(n, type) {
            let table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
            table = document.getElementById("statsTable");
            switching = true;
            dir = "asc";

            // Remove all arrow indicators
            document.querySelectorAll('th span').forEach(span => span.textContent = '');

            while (switching) {
                switching = false;
                rows = table.rows;
                for (i = 1; i < (rows.length - 1); i++) {
                    shouldSwitch = false;
                    x = rows[i].getElementsByTagName("TD")[n];
                    y = rows[i + 1].getElementsByTagName("TD")[n];

                    if (type === 'num') {
                        if ((dir == "asc" && parseFloat(x.innerHTML) > parseFloat(y.innerHTML)) || 
                            (dir == "desc" && parseFloat(x.innerHTML) < parseFloat(y.innerHTML))) {
                            shouldSwitch = true;
                            break;
                        }
                    } else if (type === 'str') {
                        if ((dir == "asc" && x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) || 
                            (dir == "desc" && x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase())) {
                            shouldSwitch = true;
                            break;
                        }
                    }
                }
                if (shouldSwitch) {
                    rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                    switching = true;
                    switchcount++;
                } else {
                    if (switchcount == 0 && dir == "asc") {
                        dir = "desc";
                        switching = true;
                    }
                }
            }

            // Add arrow indicators after sorting
            const arrowSymbol = dir == "asc" ? "&#x25BC;" : "&#x25B2;";
            document.querySelectorAll("th span")[n].innerHTML = arrowSymbol;

            renderTablePage(currentPage);
        }

        renderTablePage(currentPage);
        </script>
        """

        # Add dropdown for selecting terms and showing the corresponding map if generate_map is True
        dropdown_html = ""
        if generate_map:
            dropdown_html = f"""
            <div>
                <h2>Worldwide Search Interest for <span id="selectedTerm">{stats[0][0]}</span></h2>
                <select id="termDropdown" onchange="changeMap(this.value)">
                    {''.join([f'<option value="{term}">{term}</option>' for term in search_terms])}
                </select>
            </div>
            """

        full_html = f"""
        <html>
        <head>
        <title>Google Trends Comparison</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 20px;
            }}
            h2 {{
                text-align: center;
                margin-top: 40px;
            }}
            .container {{
                max-width: 1000px;
                margin: auto;
            }}
            .graph-container {{
                margin-bottom: 40px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
            }}
            th {{
                cursor: pointer;
                padding: 10px;
                background-color: #f2f2f2;
                border: 1px solid #ddd;
            }}
            td {{
                padding: 8px;
                border: 1px solid #ddd;
            }}
        </style>
        </head>
        <body>
        <div class="container">
            {missing_terms_html}
            <div class="graph-container">
                {dropdown_html}
                <div class="graph">
                    {graph_html}
                </div>
            </div>
            <h2>Stats</h2>
            <div>
                {stats_html}
            </div>
            {pagination_html}
        </div>
        </body>
        </html>
        """

        # Save the initial HTML with the highest average term's map
        filename = filename_entry.get()
        if not filename:
            filename = "google_trends_comparison.html"

        with open(filename, "w", encoding="utf-8") as f:
            f.write(full_html)

        if generate_map:
            # Save individual HTML files for each term's map
            for term in search_terms:
                term_map_html = generate_worldwide_map(data, search_terms, term)
                with open(f"{term}_map.html", "w", encoding="utf-8") as f:
                    f.write(term_map_html)

        webbrowser.open(f"file://{os.path.abspath(filename)}")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while generating the HTML output: {str(e)}")


# Function to generate the worldwide map
def generate_worldwide_map(data, search_terms, term):
    try:
        pytrends = TrendReq(hl='en-US', tz=360)
        pytrends.build_payload([term], cat=0, timeframe='today 12-m', geo='')
        df = pytrends.interest_by_region(resolution='COUNTRY', inc_low_vol=True)
        df = df.reset_index()

        fig = px.choropleth(df, locations='geoName', locationmode='country names', color=term,
                            title=f'Worldwide Search Interest for {term}',
                            color_continuous_scale='Viridis')

        fig.update_layout(geo=dict(showcoastlines=True, coastlinecolor="Black"))

        return pio.to_html(fig, full_html=False)
    except Exception as e:
        print(f"Error generating worldwide map: {str(e)}")
        return ""

# Function to calculate averages for selected regions
def get_selected_region_averages(region_data, term):
    if region_data.empty or term not in region_data.columns:
        return 0, 0, 0
    year_avg = region_data[term].mean()
    month_avg = region_data[term].tail(30).mean()  # Approximate month average (last 30 days)
    week_avg = region_data[term].tail(7).mean()  # Approximate week average (last 7 days)
    return year_avg, month_avg, week_avg

# Function to get comparison data for the top region
def get_top_region_comparison_data(term):
    try:
        pytrends = TrendReq(hl='en-US', tz=360)
        pytrends.build_payload([term], cat=0, timeframe='today 12-m', geo='')
        df = pytrends.interest_by_region(resolution='COUNTRY', inc_low_vol=True)
        if df.empty:
            return "N/A", 0, 0, 0
        top_region_name = df[term].idxmax()
        top_region_year_avg = df[term].mean()
        top_region_month_avg = df[term].tail(30).mean()
        top_region_week_avg = df[term].tail(7).mean()
        return top_region_name, top_region_year_avg, top_region_month_avg, top_region_week_avg
    except Exception as e:
        return "N/A", 0, 0, 0

def compare_values(value, avg):
    if value > avg:
        return "green", "&#x25B2;"  # Upwards arrow
    elif value < avg:
        return "red", "&#x25BC;"  # Downwards arrow
    else:
        return "black", ""  # No change

# Function to load search terms from a file
def load_terms_from_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("CSV Files", "*.csv")])
    if file_path:
        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()
            terms = content.split(",")
            terms = [term.strip() for term in terms if term.strip()]
            entry.delete(0, tk.END)  # Clear current entry
            entry.insert(0, ",".join(terms))  # Insert terms into the entry box

# Function to merge multiple HTML files

# Function to merge multiple HTML files
def merge_html_files():
    files = filedialog.askopenfilenames(filetypes=[("HTML Files", "*.html")])
    if not files:
        return

    merged_html = ""

    for i, file in enumerate(files):
        with open(file, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f, 'html.parser')

            # Extract the body content
            body_content = soup.find('body')
            if body_content:
                merged_html += str(body_content)
    
    # Combine everything into a single HTML structure
    final_html = f"""
    <html>
    <head>
        <title>Google Trends Merged Comparison</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 20px;
            }}
            h2 {{
                text-align: center;
                margin-top: 40px;
            }}
            .container {{
                max-width: 1000px;
                margin: auto;
            }}
            .graph-container {{
                margin-bottom: 40px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
            }}
            th, td {{
                padding: 10px;
                border: 1px solid #ddd;
            }}
        </style>
    </head>
    <body>
        {merged_html}
    </body>
    </html>
    """

    # Save the combined HTML file
    save_path = filedialog.asksaveasfilename(
        defaultextension=".html",
        initialfile="google_trends_merged",
        filetypes=[("HTML Files", "*.html")]
    )
    if save_path:
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(final_html)
        messagebox.showinfo("Success", f"Files merged and saved successfully as {os.path.basename(save_path)}!")


    # Remove duplicates in stats and sort combined stats
    combined_stats = pd.DataFrame(combined_stats)
    combined_stats = combined_stats.drop_duplicates().sort_values(by=1, ascending=False)

    # Generate the combined graph and table
    generate_combined_html(all_data, combined_stats)

def generate_combined_html(data, stats):
    # Create the combined graph
    fig = go.Figure()

    for column in data.columns:
        fig.add_trace(go.Scatter(x=data.index, y=data[column], mode='lines+markers', name=column,
                                 hoverinfo='name+y', marker=dict(size=8), line=dict(width=2)))

    fig.update_layout(
        title='Google Trends Combined Comparison',
        xaxis_title='Date',
        yaxis_title='Search Interest',
        hovermode='x unified',
        margin=dict(l=40, r=40, t=40, b=40),
        autosize=True,
        showlegend=True,
        hoverlabel=dict(bgcolor='white', bordercolor='black', font_size=12, namelength=0),
        dragmode=False,
    )

    graph_html = pio.to_html(fig, full_html=False)

    # Create the combined stats table
    stats_html = "<table id='statsTable' style='width: 100%; border-collapse: collapse;'>"
    stats_html += "<thead><tr><th>Term</th><th>Latest Value</th><th>Year Avg</th><th>Month Avg</th><th>Week Avg</th></tr></thead><tbody>"
    
    for index, row in stats.iterrows():
        stats_html += f"<tr><td>{row[0]}</td><td>{row[1]}</td><td>{row[2]}</td><td>{row[3]}</td><td>{row[4]}</td></tr>"
    
    stats_html += "</tbody></table>"

    # Combine the graph and stats into a single HTML
    full_html = f"""
    <html>
    <head>
    <title>Google Trends Combined Comparison</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
        }}
        h2 {{
            text-align: center;
            margin-top: 40px;
        }}
        .container {{
            max-width: 1000px;
            margin: auto;
        }}
        .graph-container {{
            margin-bottom: 40px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th, td {{
            padding: 10px;
            border: 1px solid #ddd;
        }}
    </style>
    </head>
    <body>
    <div class="container">
        <div class="graph-container">
            <div class="graph">
                {graph_html}
            </div>
        </div>
        <h2>Combined Stats</h2>
        <div>
            {stats_html}
        </div>
    </div>
    </body>
    </html>
    """

    # Save the combined HTML file
    save_path = filedialog.asksaveasfilename(
        defaultextension=".html",
        initialfile="google_comparisons_merged",
        filetypes=[("HTML Files", "*.html")]
    )
    if save_path:
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(full_html)
        messagebox.showinfo("Success", f"Files merged and saved successfully as {os.path.basename(save_path)}!")


    # Remove duplicates and sort combined stats
    combined_stats = pd.DataFrame(combined_stats)
    combined_stats = combined_stats.drop_duplicates().sort_values(by=1, ascending=False)

    # Generate the combined graph and table
    generate_combined_html(all_data, combined_stats)

def generate_combined_html(data, stats):
    # Create the combined graph
    fig = go.Figure()

    for column in data.columns:
        fig.add_trace(go.Scatter(x=data.index, y=data[column], mode='lines+markers', name=column,
                                 hoverinfo='name+y', marker=dict(size=8), line=dict(width=2)))

    fig.update_layout(
        title='Google Trends Combined Comparison',
        xaxis_title='Date',
        yaxis_title='Search Interest',
        hovermode='x unified',
        margin=dict(l=40, r=40, t=40, b=40),
        autosize=True,
        showlegend=True,
        hoverlabel=dict(bgcolor='white', bordercolor='black', font_size=12, namelength=0),
        dragmode=False,
    )

    graph_html = pio.to_html(fig, full_html=False)

    # Create the combined stats table
    stats_html = "<table id='statsTable' style='width: 100%; border-collapse: collapse;'>"
    stats_html += "<thead><tr><th>Term</th><th>Latest Value</th><th>Year Avg</th><th>Month Avg</th><th>Week Avg</th></tr></thead><tbody>"
    
    for index, row in stats.iterrows():
        stats_html += f"<tr><td>{row[0]}</td><td>{row[1]}</td><td>{row[2]}</td><td>{row[3]}</td><td>{row[4]}</td></tr>"
    
    stats_html += "</tbody></table>"

    # Combine the graph and stats into a single HTML
    full_html = f"""
    <html>
    <head>
    <title>Google Trends Combined Comparison</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
        }}
        h2 {{
            text-align: center;
            margin-top: 40px;
        }}
        .container {{
            max-width: 1000px;
            margin: auto;
        }}
        .graph-container {{
            margin-bottom: 40px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th, td {{
            padding: 10px;
            border: 1px solid #ddd;
        }}
    </style>
    </head>
    <body>
    <div class="container">
        <div class="graph-container">
            <div class="graph">
                {graph_html}
            </div>
        </div>
        <h2>Combined Stats</h2>
        <div>
            {stats_html}
        </div>
    </div>
    </body>
    </html>
    """

    # Save the combined HTML file
    save_path = filedialog.asksaveasfilename(
        defaultextension=".html",
        initialfile="google_comparisons_merged",
        filetypes=[("HTML Files", "*.html")]
    )
    if save_path:
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(full_html)
        messagebox.showinfo("Success", f"Files merged and saved successfully as {os.path.basename(save_path)}!")

# Set up the GUI with ttk widgets and custom styles
root = tk.Tk()
root.title("Google Trends Comparison")
root.geometry("600x500")
root.configure(bg="#f0f0f0")

# Define style for buttons and entries
style = ttk.Style()
style.configure("TButton", font=("Arial", 12), padding=10)
style.configure("TEntry", font=("Arial", 12), padding=5)

# Frame to hold entry and load button
entry_frame = ttk.Frame(root)
entry_frame.pack(pady=10)

# Entry widget to input search terms
entry = ttk.Entry(entry_frame, width=50)
entry.grid(row=0, column=0, padx=(0, 10))

# Button to load terms from file, integrated with the entry box
file_button = ttk.Button(entry_frame, text="Load from file (.csv .txt)", command=load_terms_from_file)
file_button.grid(row=0, column=1)

# Dropdown menus for selecting regions
regions = ["", "US", "CA", "GB", "AU", "DE", "FR", "IN", "JP", "CN", "BR", "ZA"]  # Add more regions as needed

region1_label = tk.Label(root, text="Select Region 1 (optional):", bg="#f0f0f0")
region1_label.pack(pady=5)
region1_dropdown = ttk.Combobox(root, values=regions, state="readonly")
region1_dropdown.pack(pady=5)

region2_label = tk.Label(root, text="Select Region 2 (optional):", bg="#f0f0f0")
region2_label.pack(pady=5)
region2_dropdown = ttk.Combobox(root, values=regions, state="readonly")
region2_dropdown.pack(pady=5)

# Option to generate the map
map_var = tk.BooleanVar(value=False)
map_checkbox = ttk.Checkbutton(root, text="Generate World Map", variable=map_var)
map_checkbox.pack(pady=5)

# Button to trigger the plotting
plot_button = ttk.Button(root, text="Plot Trends", command=lambda: plot_trends(entry.get().split(','), region1_dropdown.get(), region2_dropdown.get(), map_var.get()))
plot_button.pack(pady=10)

# Entry widget to input filename
filename_label = tk.Label(root, text="Enter filename (optional):", bg="#f0f0f0")
filename_label.pack(pady=10)

filename_entry = ttk.Entry(root, width=50)
filename_entry.pack(pady=10)

# Frame to hold the merge button
button_frame = ttk.Frame(root)
button_frame.pack(pady=20)  # Increase the padding around the frame

# Button to merge datasets with increased padding and width
merge_button = ttk.Button(button_frame, text="Merge Datasets", command=merge_html_files, width=20)
merge_button.pack(pady=10, padx=10)  # Increase both vertical and horizontal padding

# Run the GUI main loop
root.mainloop()

