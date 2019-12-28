#!usr/bin/env python
import argparse, sys

from time import sleep
from weaver import weave_reports, fetch_interfaces # TODO: Move _get_rep_type, fetch_interfaces


def main(args):
    """
    Parses input args and calls modules to perform
    the primary functions of the program, namely,
    generating PCB simulation reports based on the input conf_path
    """
    # Process input from positional args    
    conf_path = args.conf_tools 

    # Process input from optional args
    img_dir = args.image_dir # TODO: Conditional logic for image_dir opt
    sim_dir = args.simulation_dir

    # Get filename of confirmation report and grab report type 
    # TODO: Keep the directory for navigating and fetching other files
    # rep_type = _get_rep_type(conf_path)
    interfaces = fetch_interfaces(args.simulation_dir) if sim_dir else []

    # Make reports based on inputs and print confirmation
    weave_reports(conf_path)
    # print(f"Weaving of report(s) for simulation type {} complete.\n")
    
    # Close program
    sleep(1)
    _ = input("Press any key to quit.\n")

    # Success
    sys.exit(0)


if __name__ == "__main__":
    desc = """
            Weaver.py takes paths to: 
                (1) a confirmation tools report,
                (2) a Simulaton directory
                (3) a templates.txt listing templates to be used (optional)
                (4) a destination directory (optional)

            to automatically generate simulation reports 
            according to a templates listed in (3).

            For more information refer to the README.
           """
    parser = argparse.ArgumentParser(description=desc)

    # Positional args
    parser.add_argument("conf_tools", help="Path to confirmation tools for simulation reports")

    # Optional args
    parser.add_argument("-s", "--simulation_dir", nargs=1, help="Path to simulation directory") 
    parser.add_argument("-i", "--image_dir", nargs=1, help="Path to directory of images to be included in the report(s)")

    # Retrieve args
    args = parser.parse_args()   

    main(args)

