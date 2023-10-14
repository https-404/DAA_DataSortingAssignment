import pandas as pd
import time

# Insertion sort implementation
def insertion_sort(arr, descending=True):
    for i in range(1, len(arr)):
        key = arr[i]
        j = i - 1
        while j >= 0 and (arr[j] < key if descending else arr[j] > key):
            arr[j + 1] = arr[j]
            j -= 1
        arr[j + 1] = key

# Record the start time
start_time = time.time()

# Load the dataset from an Excel file
data = pd.read_excel("Dataset.xlsx")

# Extract the "Annual Income" column as a list
annual_income_column = data["Annual Income"].tolist()

# Apply the insertion sort algorithm to the "Annual Income" column in descending order
insertion_sort(annual_income_column, descending=True)

# Update the "Annual Income" column with the sorted values
data["Annual Income"] = annual_income_column

# Save the sorted data frame back to an Excel file
sorted_data = data
sorted_data.to_excel("SortedDataset_InsertionSort_Descending.xlsx", index=False)

# Record the end time
end_time = time.time()

# Calculate and display the execution time
execution_time = end_time - start_time
print(f"Execution time: {execution_time} seconds")