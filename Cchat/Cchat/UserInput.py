from Inventory import Items
from 
# Defining Max failed attempts variable
MAX_FAILED_ATTEMPTS = 2

def take_order():
       # Creating a list to store the user's order and a variable to track failed attempts
    order = []           
    failed_attempts = 0    
    # Loop to continuously prompt the user for their order until they either complete it or exceed the maximum number of failed attempts
    while True:
        # Prompting the user for their order and formatting the input for captialization 
        item = input("What would you like? ").strip().title()
        # Checking if the item is available in the shop's inventory and if it is sold out
        if item not in Items:
            print(f"Sorry, we don't sell {item}.")
            failed_attempts += 1
        elif not  Items [item]:
            print(f"Sorry, we've sold out of {item}.")
            failed_attempts += 1
            # If the item is available and not sold out, it is added to the order list and the failed attempts counter is reset
        else:
            order.append(item)                   
            failed_attempts = 0                  
            # Printing the current order to the user and asking if they would like to add anything else .join () method is used to join the items in the order list into a single string, separated by commas
            print(f"{item} - great choice! So far: {', '.join(order)}")

            if input("Anything else? (yes/no) ").strip().lower() == "yes":
                continue                         # back to the top of the loop
            return order                         # finished: hand back everything

        # only reached after a WRONG item
        if failed_attempts >= MAX_FAILED_ATTEMPTS:
            print("Let me put you through to a member of staff.")
            return None


        if input("Would you like to try something else? (yes/no) ").strip().lower() != "yes":
            return order

# Main block to call the take_order function when the script is run directly
if __name__ == "__main__":
    order = take_order()
    if order is None:
        print("(Call transferred to staff)")
    elif not order:
        print("No order taken.")
    else:
        print(f"Final order: {', '.join(order)}")