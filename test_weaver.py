import weaver
import pytest
import os

from weaver import ConfirmationTools

""" 
The following tests follow the four-phase
(1) Setup, (2) Execute, (3) Verify, (4) Teardown 
approach described by Gerard Meszaros 
in "xUnit test patterns: Refactoring test code" (2007).

Areas where (4) is left blank assumes Python's garbage collector
to be at work in freeing memory.
"""

# \\\\\\\\\\\\\\\\\\\\\\
#  HELPER FUNCTION TESTS
# //////////////////////

def set_test_params(path):
    """
    Returns dict of test parameters extracted from textfile
    """

    params = {
        "sim_dir": "",
        "si_path": "",
        "conf_path": "",
        "reviewers": ""
    }

    with open(path, "r") as f:
        for line in f.readlines():
            param_name, param_val = line.split("=")
            for k in params.keys():
                if param_name == k:
                    params[k] = param_val.rstrip()
    
    return params


PARAMS = set_test_params(os.getenv("PARAMS_PATH"))


def test_fetch_interfaces():

    # (1) Setup
    print(PARAMS)
    sim_dir = PARAMS["sim_dir"] 
    expected_if_list = ["RGMII"]

    # (2) Execute
    actual_if_list = weaver.fetch_interfaces(sim_dir)

    # (3) Verify
    assert isinstance(actual_if_list, list)
    assert actual_if_list == expected_if_list 

    # (4) Teardown
    

def test_load_template_paths():

    # (1) Setup
    expected_si_path = PARAMS["si_path"]
    expected_num_temps = 4

    # (2) Execute
    actual_templates = weaver._load_template_paths(weaver.TXT_PATH)

    # (3) Verify
    assert isinstance(actual_templates, dict)
    assert actual_templates["si"] == expected_si_path
    assert len(actual_templates) ==  expected_num_temps

    # All other values in dict are ""
    for k in actual_templates.keys():
        if not k == "si":
            assert not actual_templates[k] 

    # (4) Teardown


# \\\\\\\\\\\\\\\\\\\\\\
#  CLASS INSTANCE TESTS
# //////////////////////

# Global instance of ConfirmationTools
test_ct = ConfirmationTools(PARAMS["conf_path"])

def test_get_creators():
    
    # (1) Setup
    expected_reviewers = PARAMS["reviewers"]

    # (2) Execute
    actual_creators = test_ct.get_creators()

    # (3) Verify
    assert isinstance(actual_creators, dict)
    assert actual_creators["reviewers"] == expected_reviewers
    assert not actual_creators["preparers"] == expected_reviewers
    assert actual_creators["preparers"].find(",") > -1

    # (4) Teardown


def test_get_toc():

    # (1) Setup
    # None

    # (2) Execute
    actual_toc = test_ct.get_toc()

    # (3) Verify
    assert len(actual_toc.keys()) == 3
    for slide_nums in actual_toc.values():
        # All values have been found
        assert not slide_nums == "" and isinstance(slide_nums, list)
        for sn in slide_nums:
            assert isinstance(sn, int)

    # (4) Teardown


def test_init_reports():

    # (1) Setup
    rep_type = "si"
    interfaces = ["RGMII"]
    rep_str = f"{rep_type.upper()} Report for {interfaces[0]}"

    # (2) Execute
    reports = weaver.init_reports(rep_type, test_ct, interfaces)
    test_rep = reports[0]

    # (3) Verify
    assert len(reports) == 1
    assert test_rep.report_type[:].lower() == rep_type
    assert test_rep.pptx # Check if Presentation obj created
    assert str(test_rep) == rep_str

    # (4) Teardown