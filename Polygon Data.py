import pandas as pd
from math import radians, sin, cos, sqrt, atan2, asin,pi
import simplekml
import simplekml
from math import sin, cos, atan2, sqrt, radians
import math
from pykml.factory import KML_ElementMaker as KML
from simplekml import Snippet

#Batch name to generate KMLs
Batch_Number= "B4"
number_of_tickets_inside_circle = 10 #Min
kilometers_circle_to_look_inside = 0.5



# Use the Snippet object
snippet = Snippet()
# Excel to KML for ARPU Acc. function 
def excel_to_kml(excel_file, sheet_name):
    # Load the Excel sheet into a pandas dataframe

    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Create a new KML object
    kml = simplekml.Kml()

    # Iterate over each row in the dataframe and create a placemark for each point
    for index, row in df.iterrows():
        # Create a new placemark object
        placemark = kml.newpoint(name=row['ARPU (Acc.)'])
        
        # Create a new style for the placemark
        icon_style = simplekml.Style()
        icon_style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png' # Change the icon to a circle

        # Set the color of the placemark
        color = simplekml.Color.red  # Change the color to red
        icon_style.iconstyle.color = color

        # Assign the style to the placemark
        placemark.style = icon_style
        snippet_text = "ARPU (Acc.): " +  str(row['ARPU (Acc.)']) # Text to display before the description
       
        # placemark.description = row['Dial No'] # Description field
        placemark.snippet = Snippet(snippet_text) # Add a Snippet element with the summary text

        # Set the coordinates for the placemark
        placemark.coords = [(row['Long'], row['Lat'])]

        # Add additional data as extended data to the placemark
        for column in df.columns:
            if column not in ['Lat', 'Long','ARPU (Acc.)']:
                placemark.extendeddata.newdata(column, str(row[column]))


    # Save the KML file
    kml_file = f"ARPU Acc.kml"
    kml.save(kml_file)

    print(f"KML file {kml_file} created successfully.")


def radians_to_degrees(rad):
    return rad * 180 / pi

#Get distance between two lat and longs
def get_distance(lat1, lon1, lat2, lon2):
    R = 6371.0  # Earth's radius in kilometers

    # Convert latitude and longitude values from degrees to radians
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])

    # Calculate the differences between the latitudes and longitudes
    dlat = lat2 - lat1
    dlon = lon2 - lon1

    # Apply the Haversine formula to calculate the distance
    a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))
    distance = R * c

    return distance

# Read Excel data for batches to get lat long arpu and ticket ID
Batch_df = pd.read_excel('C:\\Users\\vb\\Desktop\\Layers\\to_KML\\Complaints & Coverage 3.xlsx', sheet_name=Batch_Number)
# print(B1_df.head())
# print(B1_df.describe())

ticket_id = list(Batch_df['Ticket ID'])
Lat = list(Batch_df['Lat'])
Long = list(Batch_df['Long'])
ARPU = list(Batch_df['ARPU'])
Problem_type = list(Batch_df['Problem Type Mapped'])

distance_difference = []
for i in range(len(Lat)):
	distance_difference.append([])

for j in range(len(Lat)):
	for i in range(len(Lat)):
		distance_difference[j].append(get_distance(Lat[j],Long[j],Lat[i],Long[i]))

for k in range(len(distance_difference)):
	distance_difference[k].sort()

center_point= 0
center_point_list =[]
m_list =[]

center_point_lat =[]
center_point_long = []
center_point_ticket = []




for m in range(len(distance_difference)):
	# print(distance_difference[0][m])
	if distance_difference[m][number_of_tickets_inside_circle+1]<kilometers_circle_to_look_inside:
		# print("Center point")
		center_point=center_point+1
		center_point_list.append(distance_difference[m][11])
		m_list.append(m)
		center_point_lat.append(Lat[m])
		center_point_long.append(Long[m])
		center_point_ticket.append(ticket_id[m])



# print("-------------------------------CENTER POINTS Count-------------------------------")
# print(center_point)
# print("-------------------------------CENTER POINTS Lat-------------------------------")
# print(center_point_lat)
# print("-------------------------------CENTER POINTS Long -------------------------------")
# print(center_point_long)
# print("-------------------------------CENTER POINTS Ticket ID-------------------------------")
# print(center_point_ticket)

# Cluster points within a given radius
def cluster_points(tickets, lats, longs, radius):
    """
    Cluster points within a given radius
    """
    clusters = []
    cluster_lats = []
    cluster_longs = []
    while len(tickets) > 0:
        ticket = tickets.pop()
        lat = lats.pop()
        lon = longs.pop()
        cluster = [ticket]
        i = 0
        while i < len(tickets):
            d = get_distance(lat, lon, lats[i], longs[i])
            if d <= radius:
                cluster.append(tickets.pop(i))
                lats.pop(i)
                longs.pop(i)
            else:
                i += 1
        clusters.append((lat, lon, cluster))
        cluster_lats.append(lat)
        cluster_longs.append(lon)
    return clusters, cluster_lats, cluster_longs

def create_kml_clusters(clusters, output_file):
    """
    Create a KML file for the clusters
    """
    kml = simplekml.Kml()
    for lat, lon, cluster in clusters:
        placemark = kml.newpoint(name=f"Cluster {lat}, {lon}", coords=[(lon, lat)])
        placemark.description = f"Tickets: {', '.join(str(t) for t in cluster)}"
    kml.save(output_file)

# Create cluster with 3 KM circle width
cluster ,cluster_lats , cluster_longs= cluster_points(center_point_ticket,center_point_lat,center_point_long,kilometers_circle_to_look_inside)


# print("-------------------------------CENTER POINTS Cluster-------------------------------")
# print((cluster))
# print("\n\n\n\n\n")
# print(cluster_lats)
# print(cluster_longs)
output_file = "clusters.kml"

# draw circle and generate KML for the clusters
# def create_kml_circle_polygon(latitudes, longitudes, output_file):
#     kml = simplekml.Kml()
#     for lat, lon in zip(latitudes, longitudes):
#         circle_points = []
#         for i in range(0, 361, 1):
#             lat_circle = math.sin(math.radians(i)) * 0.05 / math.cos(math.radians(lat)) + lat
#             lon_circle = math.cos(math.radians(i)) * 0.05 + lon
#             circle_points.append((lon_circle, lat_circle))
#         circle = kml.newpolygon(name=f"Circle around ({lat}, {lon})", outerboundaryis=circle_points)
#         circle.style.polystyle.color = simplekml.Color.rgb(255, 0, 0, 50)
#         circle.style.linestyle.color = simplekml.Color.rgb(255, 0, 0, 100)
#         circle.style.linestyle.width = 3
#     kml.save(output_file)


def create_kml_circle_polygon(latitudes, longitudes, output_file, radius_km):
    kml = simplekml.Kml()
    for lat, lon in zip(latitudes, longitudes):
        circle_points = []
        radius_deg_lat = radius_km / 111.32  # radius in degrees of latitude
        radius_deg_lon = radius_km / (111.32 * math.cos(math.radians(lat)))  # radius in degrees of longitude
        for i in range(0, 361, 1):
            lat_circle = math.sin(math.radians(i)) * radius_deg_lat + lat
            lon_circle = math.cos(math.radians(i)) * radius_deg_lon + lon
            circle_points.append((lon_circle, lat_circle))
        circle = kml.newpolygon(name=f"Circle around ({lat}, {lon})", outerboundaryis=circle_points)
        circle.style.polystyle.color = simplekml.Color.rgb(255, 0, 0, 50)
        circle.style.linestyle.color = simplekml.Color.rgb(255, 0, 0, 100)
        circle.style.linestyle.width = 3
    kml.save(output_file)

create_kml_circle_polygon(cluster_lats,cluster_longs,output_file,kilometers_circle_to_look_inside)

print("-------------- Cluster Locations-------------------")
print(cluster_lats[0])
print(cluster_longs[0])
print("-------------- Main Locations-------------------")
# print(Lat)
# print(Long)
dist_diffrence_ARPU = []

cluster_ARPU = []
cluster_problem_type =[]
for i in range(len(cluster_lats)):
    cluster_ARPU.append([])
    cluster_problem_type.append([])



for v in range(len(cluster_lats)):
    for i in range(len(Lat)):
        x = get_distance(cluster_lats[v],cluster_longs[v],Lat[i],Long[i])
        if x<kilometers_circle_to_look_inside:
            dist_diffrence_ARPU.append(x)
            cluster_ARPU[v].append(ARPU[i])
            cluster_problem_type[v].append(Problem_type[i])

# print('cluster_problem_type[0]')
# print(cluster_problem_type[0])
# print(len(cluster_problem_type[0]))

Voice_count_list = []
Data_count_list = []
VoiceData_count_list = []

for i in range(len(cluster_problem_type)):
    Voice_count_list.append([])
    Data_count_list.append([])
    VoiceData_count_list.append([])

voice_counter = 0
data_counter = 0
voiceANDdata_counter = 0

# Count each cluster voice , data, voiceNdata
# for f in range(len(cluster_lats))

# print(len(cluster_problem_type))
# print(len(cluster_lats))
for t in range(len(cluster_problem_type)):  
    for i in range(len(cluster_problem_type[t])):
        if cluster_problem_type[t][i] =="Voice":
            voice_counter+=1
        elif cluster_problem_type[t][i] =="Data":
            data_counter+=1
        elif cluster_problem_type[t][i] =="Voice/Data":
            voiceANDdata_counter+=1

    Voice_count_list[t].append(voice_counter)
    Data_count_list[t].append(data_counter)
    VoiceData_count_list[t].append(voiceANDdata_counter)
    voice_counter = 0
    data_counter = 0
    voiceANDdata_counter = 0

# List of Voice-data-both count for each cluster 
# Change list of lists to list of items
Voice_count_list = [item for sublist in Voice_count_list for item in sublist]
Data_count_list = [item for sublist in Data_count_list for item in sublist]
VoiceData_count_list = [item for sublist in VoiceData_count_list for item in sublist]


print ("voice_counter : "+  str(Voice_count_list))
print ("data_counter : " + str(Data_count_list))
print ("voice & data_counter : " +str(VoiceData_count_list))



def sum_lists(list_of_lists):
    """Returns a list of sums of the inner lists"""
    return [sum(inner_list) for inner_list in list_of_lists]



cluster_ARPU_accumlitive = sum_lists(cluster_ARPU)
print(cluster_ARPU_accumlitive)



df = pd.DataFrame({'Lat': cluster_lats, 'Long': cluster_longs, 'ARPU (Acc.)': cluster_ARPU_accumlitive , 'Voice Count': Voice_count_list, 'Data Count':Data_count_list,'Voice & Data Count':VoiceData_count_list})
df.to_excel('Center Points.xlsx', index=False)

# Load the data from the Excel file
df = pd.read_excel('Center Points.xlsx')

# '''''''''''''''''''''''''''''''''''''''

# Reading data from Center Points EXCEL
excel_file = "Center Points.xlsx"
sheet_name = "Sheet1"
excel_to_kml(excel_file, sheet_name)