from new_Account import *
from depositing import *
from withdrawing import *
from Acc_activity import *
from Acc_close import *
from balance_Inquiry import *
from listing import *
from modify import *
init(autoreset=True)

global f_name, m_name, l_name, dob, age, gender, email, pn, emp_status, ini_dep, acc_type, acc_pin, acc_status, acc_no, T_C, acc_row, ac_row


def end():
    print("Bye.....")
    print("Logging out from the server", end="")
    i = 0.5
    while i != random.randint(2, 4):
        print(".", end="")
        time.sleep(i)
        i += 0.5
    print("\nLogged out..")
    try:
        ws.save("Account Database.xlsx")
        ws.close()
        wk.save("Country-codes.xlsx")
        wk.close()
        exit()
    except Exception as e:
        print("Error!!, ", e)
        exit()


def selection():
    choice = input()
    implement = {"0": end, "1": new_Acc, "2": deposit, "3": withdraw, "4": bal_Inq, "5": list_Acc, "6": close_Acc, "7": modify_Acc}
    if choice in implement:
        implement[choice]()
    else:
        invalid = "\nPlease enter value based on given options.."
        print(Style.DIM + Fore.RED + invalid.upper())
        options()
    return choice


def options():
    print(Style.BRIGHT + Fore.BLUE + "Welcome to the Swiss Bank"+Style.BRIGHT + Fore.MAGENTA + "\nHere is the menu containing the different options:")
    print(Style.BRIGHT + Fore.MAGENTA + "\t1) New Account\n\t2) Deposit amount\n\t3) Withdraw amount\n\t4) Balance Enquiry\n\t5) All Account Holder list\n\t6) Close an account\n\t7) Modify An account\n\t0) Exit")
    print(Style.BRIGHT + Fore.GREEN + "Please choose any one option from above: " + Style.RESET_ALL, end="")
    return selection()


if __name__ == '__main__':
    while True:
        try:
            options()
        except Exception as E:
            print(Style.DIM + Fore.RED + "Error!!", E)
        finally:
            options()
# Also, when it is connected to live server, then add the column in Excel which shows the total time of account since opened, if account status == inactive, the time will not run.
