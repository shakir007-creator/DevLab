import sys
from subprocess import run, CalledProcessError
import time

# Start timing
start_time = time.time()

def ask_user_for_task(task_name):
    while True:
        try:
            user_input = input(f"\nDo you want to {task_name} (Y/N)?: ").strip().lower()
            if user_input in ['y', 'yes']:
                return True
            elif user_input in ['n', 'no']:
                return False
            else:
                raise ValueError("\nInvalid input! Please enter 'Y' or 'N'.\n")
        except ValueError as e:
            print(e)
            sys.exit(1)

def main():
    try:
        # Ask for the first task: Extract Data
        if ask_user_for_task("Extract Data from Outlook"):
            try:
                run(["python", "CKT_ID_Extractor.py"], check=True)
            except CalledProcessError as e:
                print(f"\nError executing CKT_ID_Extractor.py: {e}\n")
        else:
            print("\nSkipping data extraction.\n")

        # Ask for the second task: Insert Formula
        if ask_user_for_task("insert formulas into the sheet"):
            try:
                run(["python", "Formula_Inserter.py"], check=True)
            except CalledProcessError as e:
                print(f"\nError executing Formula_Inserter.py: {e}\n")
        else:
            print("\nSkipping formula insertion.\n")

        # Ask for the third task: Clean Data
        if ask_user_for_task("clean the data"):
            try:
                run(["python", "Data_Cleaner.py"], check=True)
            except CalledProcessError as e:
                print(f"\nError executing Data_Cleaner.py: {e}\n")
        else:
            print("\nSkipping data cleaning.\n")

    except Exception as e:
        print(f"\nAn error occurred during task execution: {e}\n")
        sys.exit(1)

    print("\nAll tasks have been completed successfully.\n")

    # End timing
    end_time = time.time()
    execution_time = end_time - start_time
    print(f"\nTotal execution time: {execution_time:.2f} seconds\n")

if __name__ == "__main__":
    main()

