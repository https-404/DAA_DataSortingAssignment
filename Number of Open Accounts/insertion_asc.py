import pandas as pd
import time

# Insertion sort implementation for ascending order
def insertion_sort(arr, column_name):
    for i in range(1, len(arr)):
        key = arr[i]
        j = i - 1
        while j >= 0 and float(arr[j][column_name]) > float(key[column_name]):
            arr[j + 1] = arr[j]
            j -= 1
        arr[j + 1] = key

# Record the start time
start_time = time.time()

# Load the dataset from an Excel file
data = pd.read_excel("Dataset.xlsx")

# Extract the data as a list of dictionaries
data_list = data.to_dict(orient='records')

# Apply the insertion sort algorithm to sort the data based on the "Number of Open Accounts" column in ascending order
insertion_sort(data_list, "Number of Open Accounts")

# Create a new DataFrame with the sorted data
sorted_data = pd.DataFrame(data_list)

# Save the sorted data frame back to an Excel file
sorted_data.to_excel("SortedDataset_InsertionSort_Ascending_NumberOfOpenAccounts.xlsx", index=False)

# Record the end time
end_time = time.time()

# Calculate and display the execution time
execution_time = end_time - start_time
print(f"Execution time: {execution_time} seconds")
