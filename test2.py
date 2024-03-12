x = [
    [1, 2, 3, 4, 5],
    [0, 2, 2, 2, 2],
    [1, 0, 2, 2, 2],
    [3, 3, 3, 3, 3]
]

# Filter out rows with the first element's value equal to 0
x = [row for row in x if row[0] != 0]

# Show the result of x
print(x)
