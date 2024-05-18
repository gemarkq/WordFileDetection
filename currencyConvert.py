from num2words import num2words

if __name__ == "__main__":
    while True:
        num = eval(input("Please input number: "))

        print(num2words(num))