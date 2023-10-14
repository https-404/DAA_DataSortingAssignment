import pandas as pd
import time

# Merge sort implementation
def merge_sort(arr, column_name, descending=False):
    if len(arr) > 1:
        mid = len(arr) // 2
        left_half = arr[:mid]
        right_half = arr[mid:]

        merge_sort(left_half, column_name, descending)
        merge_sort(right_half, column_name, descending)

        i = j = k = 0

        while i < len(left_half) and j < len(right_half):
            left_value = left_half[i][column_name]
            right_value = right_half[j][column_name]

            left_value_float = float(left_value)
            right_value_float = float(right_value)

            if (left_value_float > right_value_float) if descending else (left_value_float < right_value_float):
                arr[k] = left_half[i]
                i += 1
            else:
                arr[k] = right_half[j]
                j += 1
            k += 1

        while i < len(left_half):
            arr[k] = left_half[i]
            i += 1
            k += 1

        while j < len(right_half):
            arr[k] = right_half[j]
            j += 1
            k += 1

# Record the start time
start_time = time.time()

# Load the dataset from an Excel file
data = pd.read_excel("Dataset.xlsx")

# Extract the data as a list of dictionaries
data_list = data.to_dict(orient='records')

# Apply the merge sort algorithm to sort the data based on the "Number of Open Accounts" column in ascending order
merge_sort(data_list, "Number of Open Accounts")

# Create a new DataFrame with the sorted data
sorted_data = pd.DataFrame(data_list)

# Save the sorted data frame back to an Excel file
sorted_data.to_excel("SortedDataset_MergeSort_Ascending_NumberOfOpenAccounts.xlsx", index=False)

# Record the end time
end_time = time.time()

# Calculate and display the execution time
execution_time = end_time - start_time
print(f"Execution time: {execution_time} seconds")
