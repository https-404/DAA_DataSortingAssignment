import pandas as pd
import time

# Insertion sort implementation for descending order
def insertion_sort_descending(arr, column_name):
    for i in range(1, len(arr)):
        key = arr[i]
        j = i - 1
        while j >= 0 and float(arr[j][column_name]) < float(key[column_name]):
            arr[j + 1] = arr[j]
            j -= 1
        arr[j + 1] = key

# Record the start time
start_time = time.time()

# Load the dataset from an Excel file
data = pd.read_excel("Dataset.xlsx")

# Extract the data as a list of dictionaries
data_list = data.to_dict(orient='records')

# Calculate the median of the "Months since last delinquent" column
median = data["Months since last delinquent"].median()

# Handle "NA" values by replacing them with the median
for record in data_list:
    if record["Months since last delinquent"] == "NA":
        record["Months since last delinquent"] = str(median)

# Apply the modified insertion sort algorithm to sort the data based on the "Months since last delinquent" column in descending order
insertion_sort_descending(data_list, "Months since last delinquent")

# Create a new DataFrame with the sorted data
sorted_data = pd.DataFrame(data_list)

# Save the sorted data frame back to an Excel file
sorted_data.to_excel("SortedDataset_InsertionSort_Descending_MonthsSinceLastDelinquent.xlsx", index=False)

# Record the end time
end_time = time.time()

# Calculate and display the execution time
execution_time = end_time - start_time
print(f"Execution time: {execution_time} seconds")
