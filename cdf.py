from collections import Counter

# 給定數據
data = [
    10, 3.57, 0.168, 0, 0.206, 0.172, 1.22, 0.059, 0.827, 0.027, 0.458, 0.635,
    19, 0.186, 2.1, 6.25, 2.07, 16, 0.067, 2.1, 14.6, 5.18, 7.79, 4.68, 0.606,
    19.7, 0.496, 0.814, 26.6, 0.594, 3.63, 2.99, 0, 9.22, 0.109, 2.11, 5.46,
    11.5, 2.97, 3.96, 7.21, 16.4, 3.68, 1.77, 0.377, 0.92, 7.15, 2.29, 4.68,
    13, 14.6
]

# 計算每個數值的出現次數
value_counts = Counter(data)

# 計算頻率百分比
total_count = len(data)
frequency_percentages = {key: (value / total_count) * 100 for key, value in value_counts.items()}

# 顯示結果
for value, percentage in frequency_percentages.items():
    print(f"數值: {value}, 頻率百分比: {percentage:.2f}%")
