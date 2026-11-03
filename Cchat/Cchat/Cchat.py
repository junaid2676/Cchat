# Cchat.py importing the date and time modules form python library
from datetime import datetime
from UserInput import take_order
from inventory import Items

# Welcome function to greet the user and take their order
def welcome():
    print("Hi there! How can I help you?")
    order = input("Please enter your order: ")

# Function to check the shop's opening and closing times
def shop_timing():
    input("Please press the enter button.")
    opening_time = datetime.strptime("10:00:00", "%H:%M:%S").time()
    closing_time = datetime.strptime("23:00:00", "%H:%M:%S").time()
    current_time = datetime.now().time()
# Checking if the current time is within the shop's opening hours
    if opening_time <= current_time < closing_time:
        # Calling the welcome function to greet the user and take their order
        welcome()                      
        # Calling the take_order function to process the user's order
        take_order()

    else:
        print("Sorry, the shop is closed. Please come back during our working hours.")


# Creating --name main-- block to call the shop_timing function when the script is run directly
if __name__ == "__main__":
    # Calling the shop_timing function to check if the shop is open or closed
    shop_timing()