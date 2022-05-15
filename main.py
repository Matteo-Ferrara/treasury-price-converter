from fractions import Fraction
from unicodedata import numeric
import sys
import xlwings as xw


@xw.func
def treasury_price_converter(price):
    """
    Convert fractional price to decimal price.
    """
    splitted_price = price.split("-")
    # If there's no fractional part return price
    if len(splitted_price) == 1:
        return int(splitted_price[0])
    # Regular HH-TTF format
    elif len(splitted_price) == 2:
        handle, thirty_seconds = int(splitted_price[0]), splitted_price[1]
    # Negative prices
    elif len(splitted_price) == 3:
        handle, thirty_seconds = int(splitted_price[1]), splitted_price[2]
    else:
        print(f"Expected price in HH-TTF format.")
        sys.exit(1)

    # Whole 32nds
    whole_points = float(thirty_seconds[0:2])
    # If there's no fractional part set fractional value to 0
    if len(thirty_seconds) == 2:
        fractional_value = 0
    else:
        # Select fractional part
        if len(thirty_seconds.split(" ")) == 1:
            fractional_part = thirty_seconds[-1]
        else:
            fractional_part = thirty_seconds.split(" ")[1]

        # + sign is equivalent to half of 1/8
        if fractional_part == "+":
            fractional_value = 0.0625
        else:
            try:
                # Integer display
                integer_display = (
                    float(fractional_part)
                    if float(fractional_part) < 5
                    else (float(fractional_part) - 1)
                )
                fractional_value = integer_display / 8

            except ValueError:
                try:
                    # Fraction as unicode
                    fractional_value = numeric(fractional_part)
                except TypeError:
                    # Fraction as str
                    fractional_value = float(Fraction(fractional_part))

    return handle + (whole_points + fractional_value) / 32


if __name__ == "__main__":
    # Tests
    print(treasury_price_converter("116-272"))
    print(treasury_price_converter("116-27¼"))
    print(treasury_price_converter("116-27 1/4"))
    print(treasury_price_converter("115-16¾"))
    print(treasury_price_converter("115-16 3/4"))
    print(treasury_price_converter("97-16¾"))
    print(treasury_price_converter("97-16 3/4"))


