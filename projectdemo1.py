import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import webbrowser
import tkinter as tk
from tkinter import ttk


# Function to open Google Form
def open_google_form():
    google_form_url = "https://docs.google.com/forms/d/1JR-dCZZFvGHvwoaIEDXZF4FgINv7BXsgrf0R9mpaqs0/edit"
    webbrowser.open(google_form_url)


# Functions for City-specific Analysis
def top_5_profitable_cars_in_city(city_data):
    top_5_profitable = city_data.nlargest(5, 'Profit/Loss')
    print("Top 5 Profitable Cars in the selected city: \n", top_5_profitable)


def top_5_loss_making_cars_in_city(city_data):
    top_5_loss_making = city_data.nsmallest(5, 'Profit/Loss')
    print("Top 5 Loss-making Cars in the selected city: \n", top_5_loss_making)


def highest_mileage_car_in_city(city_data):
    highest_mileage = city_data.nlargest(1, 'Mileage (km/L)')
    print("Car with the Highest Mileage in the selected city: \n", highest_mileage)


def city_profit_loss_bar_plot(city_data):
    plt.bar(city_data['Car No'], city_data['Profit/Loss'])
    plt.title('Profit/Loss of Cars in the Selected City')
    plt.xlabel('Car No')
    plt.ylabel('Profit/Loss')
    plt.show()


# Function to handle city-wise analysis after selecting a city
def city_wise_analysis(city):
    df = pd.read_excel('Data analytic_power_bi_project.xlsx')
    city_data = df[df['City'] == city]

    if city_data.empty:
        print(f"No data available for {city}.")
        return

    # Clear previous widgets if any
    for widget in root.pack_slaves():
        if isinstance(widget, tk.Button) or isinstance(widget, tk.Label):
            widget.pack_forget()

    # Create individual buttons for each type of analysis in the selected city
    create_hover_button(root, text="Top 5 Profitable Cars", command=lambda: top_5_profitable_cars_in_city(city_data)).pack(pady=10)
    create_hover_button(root, text="Top 5 Loss-making Cars", command=lambda: top_5_loss_making_cars_in_city(city_data)).pack(pady=10)
    create_hover_button(root, text="Highest Mileage Car", command=lambda: highest_mileage_car_in_city(city_data)).pack(pady=10)
    create_hover_button(root, text="Show Profit/Loss Bar Plot", command=lambda: city_profit_loss_bar_plot(city_data)).pack(pady=10)


# Function to show city-wise analysis after selecting a city from the dropdown
def select_city_for_analysis():
    city = city_dropdown.get()
    city_wise_analysis(city)


# Overall Analysis Functions
def overall_profit_loss_by_city(df):
    city_profit_loss = df.groupby('City')['Profit/Loss'].sum()
    print("Overall Profit/Loss by City:\n", city_profit_loss)

    # Bar plot for overall profit/loss by city
    plt.bar(city_profit_loss.index, city_profit_loss.values)
    plt.title('Overall Profit/Loss by City')
    plt.xlabel('City')
    plt.ylabel('Total Profit/Loss')
    plt.ylim(200000, 300000)
    plt.xticks(rotation=45)  # Rotate city names for better readability
    plt.show()


def overall_mileage_analysis(df):
    average_mileage = df.groupby("City")['Mileage (km/L)'].mean()
    print(f"Overall Average Mileage: {average_mileage} km/L")
    plt.bar(average_mileage.index, average_mileage.values)
    plt.title("Overall Milage Analysis")
    plt.xlabel('City')
    plt.ylabel('Total Profit/Loss')
    plt.ylim(12.5, 20)
    plt.xticks(rotation=45)  # Rotate city names for better readability
    plt.show()


def overall_top_5_mileage_cars(df):
    top_5_mileage = df.nlargest(5, 'Mileage (km/L)')
    print("Top 5 Mileage Cars Overall:\n", top_5_mileage)


def overall_bottom_5_mileage_cars(df):
    bottom_5_mileage = df.nsmallest(5, 'Mileage (km/L)')
    print("Bottom 5 Mileage Cars Overall:\n", bottom_5_mileage)


# Function to show overall analysis options after clicking the "Overall Analysis" button
def overall_analysis_options():
    create_hover_button(root, text="Overall Profit/Loss Analysis by City", command=lambda: overall_profit_loss_by_city(df)).pack(pady=10)
    create_hover_button(root, text="Overall Mileage Analysis", command=lambda: overall_mileage_analysis(df)).pack(pady=10)
    create_hover_button(root, text="Overall Top 5 Mileage Cars", command=lambda: overall_top_5_mileage_cars(df)).pack(pady=10)
    create_hover_button(root, text="Overall Bottom 5 Mileage Cars", command=lambda: overall_bottom_5_mileage_cars(df)).pack(pady=10)


# Function to show the "Overall Analysis" button and city-wise analysis options
def show_analysis():
    button1.pack_forget()
    button2.pack_forget()

    # Button for overall analysis
    create_hover_button(root, text="Overall Analysis", command=overall_analysis_options).pack(pady=10)

    # Dropdown for selecting city
    tk.Label(root, text="Select City for City-wise Analysis:").pack(pady=10)
    city_dropdown.pack(pady=10)

    # Button for city-wise analysis
    create_hover_button(root, text="City Wise Analysis", command=select_city_for_analysis).pack(pady=10)


# Function to show new buttons after the first one is clicked
def button_clicked():
    button.pack_forget()

    global button1, button2
    button1 = create_hover_button(root, text="Add More Driver Data", command=open_google_form)
    button1.pack(pady=10)

    button2 = create_hover_button(root, text="Show Analysis", command=show_analysis)
    button2.pack(pady=10)


# Function to create a button with hover effect
def create_hover_button(master, text, command):
    button = tk.Button(master, text=text, command=command, padx=10, pady=10, bg="lightgray")

    # Function to change background color on hover
    def on_enter(event):
        button['background'] = 'blue'

    def on_leave(event):
        button['background'] = 'lightgray'

    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)
    return button


# Load the data once to extract the list of cities
df = pd.read_excel('Data analytic_power_bi_project.xlsx')
city_list = df['City'].unique().tolist()

# Create the main Tkinter window
root = tk.Tk()
root.title("Data Analysis App")

# Create a dropdown for selecting city (initially hidden)
city_dropdown = ttk.Combobox(root, values=city_list)

# Create the first button to trigger the next buttons
button = create_hover_button(root, text="Car Performance Analysis", command=button_clicked)
button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
