from itertools import combinations

# 将数字粘贴进数组，用逗号分隔
numbers = []
target1 = 738664.91


def find_combinations(numbers, target):
    for r in range(1, len(numbers) + 1):
        for combo in combinations(numbers, r):
            if abs(sum(combo) - target) < 1e-2:  # 允许一定的浮点误差
                return combo
    return None

result1 = find_combinations(numbers, target1)

print("组合等于 150,934.23:", result1)