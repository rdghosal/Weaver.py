#!usr/bin/env python
import argparse, sys

from time import sleep
from weaver import weave_reports # TODO: Move _get_rep_type, fetch_interfaces


def main():
    """
    Parses input args and calls modules to perform
    the primary functions of the program, namely,
    generating PCB simulation reports based on the input conf_path
    """
    desc = """
            Weaver.py takes paths to: 
                (1) a confirmation tools report,
                (2) a Simulaton directory

            For more information refer to the README.
           """
    parser = argparse.ArgumentParser(description=desc)

    # Positional args
    parser.add_argument("conf_tools", help="Path to confirmation tools for simulation reports")

    # Optional args
    parser.add_argument("-s", "--simulation_dir", nargs=1, help="Path to simulation directory") 
    # parser.add_argument("-i", "--image_dir", nargs=1, help="Path to directory of images to be included in the report(s)")

    # Retrieve args
    args = parser.parse_args()   

    # Process input from positional args    
    conf_path = args.conf_tools 

    # Process input from optional args
    # img_dir = args.image_dir 
    sim_dir = args.simulation_dir if args.simulation_dir else ""

    # Get filename of confirmation report and grab report type 
    # TODO: Keep the directory for navigating and fetching other files
    # rep_type = _get_rep_type(conf_path)

    # Make reports based on inputs and print confirmation
    exit_code = 0
    try:
        weave_reports(conf_path, sim_dir)
    except:
        exit_code = 1
   # print(f"Weaving of report(s) for simulation type {} complete.\n")
    
    # Close program
    print()
    _ = input("Press any key to quit.")
    print(f"Weaver.py finished with Exit Code: {exit_code}")
    # Success
    sys.exit(exit_code)

# if __name__ == "__main__":
#     main()