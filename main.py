import locale
import os

from scrapper import scrap_currencies


def main():
    # To make sure the decimal separator won't mix with the thousand separator.
    locale.setlocale(locale.LC_ALL, 'en_US')

    path = "Currencies"
    is_exist = os.path.exists(path)
    if not is_exist:
        os.makedirs(path)
        print("Currencies directory is created!")

    scrap_currencies()

    print("All done. Quitting...")


if __name__ == "__main__":
    main()
