import sys
import pandas as pd
import matplotlib.pyplot as plt
import geopandas as gpd
from fpdf import FPDF
import plotly.graph_objects as go

#Import our csv file using Panda------------------------------------------------------------------------------------------
#https://www.kaggle.com/datasets/beridzeg45/tallest-buildings-in-the-world
input_dir = r"C:\Temp\py"
csv_file = input_dir + "\\tallest_buildings_in_the_world.csv"
df = pd.read_csv(csv_file, encoding='ISO-8859-1')#.head(500)
print (df.head(10))
print (list(df))

# Sort by Height and Year
df_by_height = df.sort_values(by="Height (m)", ascending=False)
df_by_year = df.head(200).sort_values(by="Completion Year")

#------------------------------------------------------------------------------------------
# Excel Export
writer = pd.ExcelWriter("Buildings_Sorted.xlsx", engine='xlsxwriter')
df_by_height.to_excel(writer, sheet_name="Buildings", index=False)

worksheet = writer.sheets["Buildings"]
worksheet.conditional_format("F2:F10000", {"type": "3_color_scale"})

#Add formula
worksheet.write(0, 11, "% Eiffel Tower")
for row_num in range(len(df)):
    formula = f'=F{row_num + 2}/{300}*100'  #python counts from 0
    worksheet.write_formula(row_num + 1, 11, formula) #excel counts from 1
percent_format = writer.book.add_format({'num_format': '0.0"%"'})
worksheet.set_column(11, 11, 20, percent_format)


# Create a bar chart
chart =  writer.book.add_chart({'type': 'column'})

# Configure chart: categories = names, values = heights
chart.add_series({
    'name':       'Buildings',
    'categories': ['Buildings', 1, 1, len(df), 1],  # column B, from row 2 to last
    'values':     ['Buildings', 1, 5, len(df), 5],  # column F, from row 2 to last
    'border':     {'color': 'blue'}
})

# Chart title and axis labels
chart.set_title({'name': 'Building Heights'})
chart.set_y_axis({'name': 'Height (m)'})
worksheet.insert_chart(4, 15, chart, {'x_offset': 25, 'y_offset': 10})
writer.close()
    
# Image export------------------------------------------------------------------------------
df = df_by_year
custom_colors = {"China": "red","United Arab Emirates": "green", "United States": "blue"}
bar_colors = df["Country"].apply(lambda x: custom_colors.get(x, "gray"))
plt.figure(figsize=(30, 20))
plt.bar(df["Name"], df["Height (m)"], color=bar_colors)
plt.axhline(y=300, color='r', linestyle='--', label="Eiffel Tower")
plt.xticks(ticks=range(len(df)), labels=df["Completion Year"],rotation=90)#rotation=45
plt.ylabel("Height (m)")
plt.title("High Buildings")

plt.savefig("building_heights.png")
plt.show()

# Map export-------------------------------------------------------------------------------
df = df_by_height
gdf = gpd.GeoDataFrame(df,geometry=gpd.points_from_xy(df.Lon, df.Lat),crs="EPSG:4326")
#https://github.com/nvkelso/natural-earth-vector/tree/master/geojson
world = gpd.read_file(input_dir + "\\ne_50m_admin_0_countries.geojson")

fig, ax = plt.subplots(figsize=(15, 10))
world.plot(ax=ax, color='gray')
gdf.plot(ax=ax, color='blue', markersize=50)
for idx, row in gdf.head(10).iterrows():
    ax.annotate(row["Name"], xy=(row["Lon"], row["Lat"]), xytext=(3, 3), textcoords="offset points")
plt.title("Building Locations")
plt.savefig("building_map.png")
plt.show()
plt.close()



#HTML export------------------------------------------------------------------------------------------
fig = go.Figure()
fig.add_trace(go.Scattergeo(
    lon=df["Lon"],
    lat=df["Lat"],
    text=df["Name"] + " (" + df["Height (m)"].astype(str) + "m)",
    marker=dict(size=10,color=df["Height (m)"])))
fig.write_html("3d_buildings_map.html")





#PDF export------------------------------------------------------------------------------------------
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=12)
pdf.cell(200, 10, "Building Heights Comparison", ln=True, align='C')
pdf.image("building_heights.png", x=10, y=20, w=180)

pdf.add_page()
pdf.cell(200, 10, "Map of Building Locations", ln=True, align='C')
pdf.image("building_map.png", x=10, y=20, w=180)
pdf.add_page()
# Title


pdf.set_font("Arial", 'B', 14)
pdf.cell(0, 10, "Top 50 Tallest Buildings", ln=True, align='C')
pdf.ln(5)

# Column headers and widths
headers = ["Name", "Height (m)", "Year Built", "City", "Country"]
col_widths = [50, 25, 25, 35, 35]

# Header row
pdf.set_font("Arial", 'B', 10)
for i, header in enumerate(headers):
    pdf.cell(col_widths[i], 8, header, border=1, align='C')
pdf.ln()

# Table rows
pdf.set_font("Arial", size=10)
for _, row in df.head(500).iterrows():
    pdf.cell(col_widths[0], 8, str(row["Name"]), border=1)
    pdf.cell(col_widths[1], 8, str(row["Height (m)"]), border=1, align='C')
    pdf.cell(col_widths[2], 8, str(row["Completion Year"]), border=1, align='C')
    pdf.cell(col_widths[3], 8, str(row["City"]), border=1)
    pdf.cell(col_widths[4], 8, str(row["Country"]), border=1)
    pdf.ln()
    
pdf.output("Building_Report.pdf")

#Use of a Function
def simple_distance(lat1, lon1, lat2, lon2):
    # Approximate conversion: 1 degree ≈ 111 km
    return ((lat1 - lat2)**2 + (lon1 - lon2)**2)**0.5 * 111


nearby_counts = []
for i, building in df.head(50).iterrows():
    count = 0
    for j, other_building in df.iterrows():
        dist = simple_distance(building["Lat"], building["Lon"], other_building["Lat"], other_building["Lon"])
        
        if dist <= 40 and dist > 0:
            count += 1
    print (building["Name"], "--", count, "buildings nearby")
