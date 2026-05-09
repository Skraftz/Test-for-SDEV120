import random

def greaterThan(x, y):
    if x > y:
        return True
    return False
while True:
    a = random.randint(-100, 100)
    b = random.randint(-150, 150)
    c = greaterThan(a, b)

    print(f"The statement {a} is greater than {b} is {c}")