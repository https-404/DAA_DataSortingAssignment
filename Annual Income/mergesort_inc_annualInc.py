import pandas as pd
import time

# Merge sort implementation
def merge_sort(arr):
    if len(arr) > 1:
        mid = len(arr) // 2
        left_half = arr[:mid]
        right_half = arr[mid:]

        merge_sort(left_half)
        merge_sort(right_half)

        i = j = k = 0

        while i < len(left_half) and j < len(right_half):
            if left_half[i] < right_half[j]:
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

# Extract the "Annual Income" column as a list
annual_income_column = data["Annual Income"].tolist()

# Apply the merge sort algorithm to the "Annual Income" column
merge_sort(annual_income_column)

# Update the "Annual Income" column with the sorted values
data["Annual Income"] = annual_income_column

# Save the sorted data frame back to an Excel file
sorted_data = data
sorted_data.to_excel("SortedDataset_MergeSort_Years of Credit History.xlsx", index=False)

# Record the end time
end_time = time.time()

# Calculate and display the execution time
execution_time = end_time - start_time
print(f"Execution time: {execution_time} seconds")
