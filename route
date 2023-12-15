import googlemaps
import pandas as pd
from itertools import permutations
from tkinter import Tk, Label, Button, filedialog, StringVar

def calculate_distance(client, origins, destinations):
    distance_matrix = client.distance_matrix(
        origins=origins,
        destinations=destinations,
        mode="driving",
    )

    distances = []
    for row in distance_matrix["rows"]:
        row_distances = [element["distance"]["value"] for element in row["elements"]]
        distances.append(row_distances)

    return distances

def generate_routes(client, orders):
    all_routes = []
    order_locations = [order["location"] for order in orders]
    distance_matrix = calculate_distance(client, order_locations, order_locations)

    for i in range(len(orders)):
        current_route = [i]
        total_distance = 0
        route_distances = []

        for j in range(len(orders) - 1):
            # Find the next destination that hasn't been visited yet
            next_destination = min(set(range(len(orders))) - set(current_route), key=lambda x: distance_matrix[current_route[-1]][x])

            # Add the next destination to the route
            current_route.append(next_destination)

            # Calculate and update distances
            segment_distance = distance_matrix[current_route[-2]][current_route[-1]]
            route_distances.append(segment_distance)
            total_distance += segment_distance

        all_routes.append((current_route, total_distance, route_distances))

    return all_routes


def find_optimal_route(client, orders):
    all_routes = generate_routes(client, orders)
    optimal_route = min(all_routes, key=lambda x: x[1])
    return optimal_route

def process_excel_file(file_path):
    # Replace 'YOUR_API_KEY' with your actual Google Maps API key.
    api_key = 'Add YOUR API KEY'
    gmaps = googlemaps.Client(key=api_key)

    df = pd.read_excel(file_path)

    df[['Latitude', 'Longitude']] = df['CF.Lat & Long'].str.split(',', expand=True)
    df['Latitude'] = pd.to_numeric(df['Latitude'])
    df['Longitude'] = pd.to_numeric(df['Longitude'])

    orders = [
        {
            "id": i + 1,
            "location": (row["Latitude"], row["Longitude"]),
            "display_name": row["display Name"]
        }
        for i, row in df.iterrows()
    ]

    optimal_route = find_optimal_route(gmaps, orders)

    optimal_route_df = pd.DataFrame({
        'OrderID': [orders[i]["id"] for i in optimal_route[0]],
        'Latitude': [orders[i]["location"][0] for i in optimal_route[0]],
        'Longitude': [orders[i]["location"][1] for i in optimal_route[0]],
        'Display Name': [orders[i]["display_name"] for i in optimal_route[0]]
    })

    optimal_route_df['Distance to Next (meters)'] = [0] + optimal_route[2]

    output_excel_file = 'optimal_route_output.xlsx'
    optimal_route_df.to_excel(output_excel_file, index=False)

    return output_excel_file

# Tkinter GUI
class App:
    def __init__(self, root):
        self.root = root
        root.title("Optimal Route Finder")

        self.label = Label(root, text="Select Excel File:")
        self.label.pack()

        self.button = Button(root, text="Browse", command=self.browse_file)
        self.button.pack()

        self.process_button = Button(root, text="Process", state="disabled", command=self.process_file)
        self.process_button.pack()

        self.result_label_text = StringVar()
        self.result_label = Label(root, textvariable=self.result_label_text)
        self.result_label.pack()

        self.open_button = Button(root, text="Open Output File", state="disabled", command=self.open_output_file)
        self.open_button.pack()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.label.config(text=f"Selected File: {file_path}")
            self.process_button.config(state="active")
            self.file_path = file_path

    def process_file(self):
        if hasattr(self, 'file_path'):
            output_file_path = process_excel_file(self.file_path)
            self.result_label_text.set(f"Optimal Route saved to: {output_file_path}")
            self.open_button.config(state="active")
            self.output_file_path = output_file_path

    def open_output_file(self):
        if hasattr(self, 'output_file_path'):
            import os
            os.system(f'start excel "{self.output_file_path}"')

# Run the GUI
root = Tk()
app = App(root)
root.mainloop()
