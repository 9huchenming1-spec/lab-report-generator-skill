"""Synthetic example code accompanying the lab-report-generator skill."""


def insertion_sort(values):
    result = list(values)
    for index in range(1, len(result)):
        current = result[index]
        position = index - 1
        while position >= 0 and result[position] > current:
            result[position + 1] = result[position]
            position -= 1
        result[position + 1] = current
    return result


def binary_search(values, target):
    left, right = 0, len(values) - 1
    while left <= right:
        middle = (left + right) // 2
        if values[middle] == target:
            return middle
        if values[middle] < target:
            left = middle + 1
        else:
            right = middle - 1
    return -1


if __name__ == "__main__":
    original = [5, 2, 4, 2]
    print("sort normal:", insertion_sort(original), "original:", original)
    print("sort empty:", insertion_sort([]))
    print("sort single:", insertion_sort([7]))
    print("search found:", binary_search([1, 3, 5, 7, 9], 7))
    print("search missing:", binary_search([1, 3, 5, 7, 9], 4))
    print("search empty:", binary_search([], 7))
