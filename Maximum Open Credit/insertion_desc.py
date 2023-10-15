import pandas as pd
import time

# Insertion sort implementation for descending order with NaN handling
def insertion_sort_descending(arr, column_name):
    for i in range(1, len(arr)):
        key = arr[i]
        j = i - 1
        while j >= 0 and (pd.isna(arr[j][column_name]) or (not pd.isna(arr[j][column_name]) and not pd.isna(key[column_name]) and arr[j][column_name] < key[column_name])):
            arr[j + 1] = arr[j]
            j -= 1
        arr[j + 1] = key

# Record the start time
start_time = time.time()

# Load the dataset from an Excel file
data = pd.read_excel("Dataset.xlsx")

# Extract the data as a list of dictionaries
data_list = data.to_dict(orient='records')

# Apply insertion sort in descending order with NaN handling to sort the data based on the "Maximum Open Credit" column
insertion_sort_descending(data_list, "Maximum Open Credit")

# Create a new DataFrame with the sorted data
sorted_data = pd.DataFrame(data_list)

# Save the sorted data frame back to an Excel file
sorted_data.to_excel("SortedDataset_InsertionSort_Descending_MaximumOpenCredit.xlsx", index=False)

# Record the end time
end_time = time.time()

# Calculate and display the execution time
execution_time = end_time - start_time
print(f"Execution time (Descending): {execution_time} seconds")
