import pandas as pd
import time

# Merge sort implementation for ascending order
def merge_sort_ascending(arr, column_name):
    if len(arr) > 1:
        mid = len(arr) // 2
        left_half = arr[:mid]
        right_half = arr[mid:]

        merge_sort_ascending(left_half, column_name)
        merge_sort_ascending(right_half, column_name)

        i = j = k = 0

        while i < len(left_half) and j < len(right_half):
            if left_half[i][column_name] < right_half[j][column_name]:
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

# Apply merge sort algorithm to sort the data based on the "Bankruptcies" column in ascending order
merge_sort_ascending(data_list, "Bankruptcies")

# Create a new DataFrame with the sorted data
sorted_data = pd.DataFrame(data_list)

# Save the sorted data frame back to an Excel file
sorted_data.to_excel("SortedDataset_MergeSort_Ascending_Bankruptcies.xlsx", index=False)

# Record the end time
end_time = time.time()

# Calculate and display the execution time
execution_time = end_time - start_time
print(f"Execution time: {execution_time} seconds")
