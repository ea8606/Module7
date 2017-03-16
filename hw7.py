"""This script displays a menu to view a list of users, add a user to the list,
   remove a user, or look up a particular user.  The list is actually a python sorted dictionary."""

from sortedcontainers import SortedDict
import sys


def print_menu():
    """Displays a menu , must enter  1,2,3,4,5..."""

    print('1. Print Users')
    print('2. Add a User')
    print('3. Remove a User')
    print('4. Lookup a User Name')
    print('5. Quit')
    print()


# Create dictionary with key = Names, value = user_name
usernames = SortedDict()
usernames['Summer'] = 'summerela'
usernames['William'] = 'GoofyFish'
usernames['Steven'] = 'LoLCat'
usernames['Zara'] = 'zanyZara'
usernames['Renato'] = 'songDude'

# setup counter to store menu choice
menu_choice = 0

# display your menu
print_menu()

# as long as the menu choice isn't "quit" get user options
while menu_choice != 5:
    try:
        # get menu choice from user
        menu_choice = int(input("Type in a number (1-5): "))
    except ValueError:
        print("Please input a number,(1,2,3,4,5,...)")
        menu_choice = 6

    # view current entries
    if menu_choice == 1:
        print("Current Users:")
        for x, y in usernames.items():
            print("Name: {} \tUser Name: {} \n".format(x, y))

    # add an entry
    elif menu_choice == 2:
        user_user_name_update = {}
        print("Add User")
        try:
            name = input("Name: ")
        except ValueError:
            print("Enter your name")
            exit()

    # Check that user entered something for a User Name

        if name == "":
            print("Must enter a name")
            exit()

        # Check if user name exists in dictionary, don't allow the change if it does

        if name in usernames.keys():
            print("User {} already exists...choose another name for user".format(name))
            exit()

        # Add a username

        try:
            username = input("User Name: ")
        except ValueError:
            print("Enter a user name for you")
            exit()
        if username == "":
            print("Must enter a username")
            exit()

        # Check if the username exists... don't allow change if it does
        if username in usernames.values():
            print("Username '{}' already exists....choose another".format(username))
            exit()

        # Update the dictionary with user and username

        user_user_name_update = {name: username}
        usernames.update(user_user_name_update)
        print("New user: {}, User Name: {}".format(name, usernames[name]))

    # remove an entry, can be user's name or user's username

    elif menu_choice == 3:
        print("Remove User")
        try:

        # Remove by user

            name = input("Name or User Name: ")
        except ValueError:
            print("Must enter a name")
            exit()
        if name == "":
            print("Must enter a name")
            exit()
        if name in usernames:
            del usernames[name]
            print("Success - User name {} deleted".format(name))

        # Remove by username too! Check to make sure that username exists first

        elif name in usernames.values():
            for x in usernames:
                if usernames[x] == name:
                    del usernames[x]
                    print("Success! Username {} deleted".format(name))
        else:
            print("Username '{}', not found".format(name))

    # view user name
    elif menu_choice == 4:
        print("Lookup User")
        try:
            name = input("Name: ")
        except ValueError:
            print("Must enter a user's name")
        finally:
            if name == "":
                print("No user name entered, exiting")
                exit()
        if name in usernames:
            print("Name: {} \tUser Name: {} \n".format(name, usernames[name]))
        else:
            print("Username not found for User:", name)

    # is user enters something strange, show them the menu
    elif menu_choice != 5:
        print_menu()