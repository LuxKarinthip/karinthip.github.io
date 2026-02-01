#หา Lat-Lon โดยใช้ ArcGIS()

from geopy.geocoders import ArcGIS

# Initialize the geolocator
nom = ArcGIS()


# Assuming the Excel file has a column named 'Address' with the addresses
addresses = df['ACCIDENT_LOCATION']

# Prepare lists to store the results
latitudes = []
longitudes = []

# Loop through each address to get the latitude and longitude
for address in addresses:
    try:
        location = nom.geocode(address)
        if location:
            latitudes.append(location.latitude)
            longitudes.append(location.longitude)
        else:
            latitudes.append(None)
            longitudes.append(None)
    except Exception as e:
        print(f"Error geocoding address {address}: {e}")
        latitudes.append(None)
        longitudes.append(None)

# Add the results to the DataFrame
df['Latitude'] = latitudes
df['Longitude'] = longitudes

# Save the updated DataFrame to a new Excel file
output_file_path = 'C:/Users/Karinthip/Desktop/Jupyter/result.xlsx'  # Replace with the desired output file path
df.to_excel(output_file_path, index=False)

print("Geocoding completed. Results saved to", output_file_path)
