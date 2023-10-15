import pandas as pd
import time
import math

# Merge sort implementation for descending order with NaN handling
def merge_sort_descending(arr, column_name):
    def custom_compare(row1, row2):
        val1 = row1[column_name]
        val2 = row2[column_name]

        if math.isnan(val1) and math.isnan(val2):
            return 0
        if math.isnan(val1):
            return 1
        if math.isnan(val2):
            return -1

        return int(val2 - val1)

    if len(arr) > 1:
        mid = len(arr) // 2
        left_half = arr[:mid]
        right_half = arr[mid:]

        merge_sort_descending(left_half, column_name)
        merge_sort_descending(right_half, column_name)

        i = j = k = 0

        while i < len(left_half) and j < len(right_half):
            if custom_compare(left_half[i], right_half[j]) <= 0:
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

# Apply merge sort in descending order with NaN handling to sort the data based on the "Maximum Open Credit" column
merge_sort_descending(data_list, "Maximum Open Credit")

# Create a new DataFrame with the sorted data
sorted_data = pd.DataFrame(data_list)

# Save the sorted data frame back to an Excel file
sorted_data.to_excel("SortedDataset_MergeSort_Descending_MaximumOpenCredit.xlsx", index=False)

# Record the end time
end_time = time.time()

# Calculate and display the execution time
execution_time = end_time - start_time
print(f"Execution time (Descending): {execution_time} seconds")
