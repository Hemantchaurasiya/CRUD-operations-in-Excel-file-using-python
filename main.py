from logging import exception
from excel import Excel
ex = Excel()

def main():
    while True:
        print("Press 1 for insert the student :")
        print("Press 2 for show all the students :")
        print("press 3 for show specific student :")
        print("Press 4 for update the student :")
        print("Press 5 for delete the student :")
        print("Press 6 for exit :")
        try:
            choice = int(input())
            if choice==1:
                userid = input("Enter UserId : ")
                username = input("Enter Name : ")
                email = input("Enter Email : ")
                phone = input("Enter Phone :")
                year = int(input("Enter Year : "))
                department = input("Enter Department : ")

                ex.insert_user(userid,username,email,phone,year,department)
                print("User inserted Successfully....")

            elif choice==2:
                ex.fetch_all_user()

            elif choice==3:
                userid = input("Enter UserId : ")
                ex.fetch_user(userid)

            elif choice==4:
                userid = input("Enter UserId : ")
                username = input("Enter Name : ")
                email = input("Enter Email : ")
                phone = input("Enter Phone :")
                year = int(input("Enter Year : "))
                department = input("Enter Department : ")
                ex.update_user(userid,username,email,phone,year,department)
                print("User updated Successfully....")

            elif choice==5:
                userid = input("Enter UserId : ")
                ex.delete_user(userid)
                print("User deleted Successfully....")

            elif choice==6:
                break
            else:
                print("Invalid key ! please try again....")
        except exception as e:
            print(e)
            print("Invalid user ! please try again....")

if __name__ == "__main__":
    main()