import weaver
import pytest
import os
import win32com.client as win32


""" 
The following tests follow the four-phase
(1) Setup, (2) Execute, (3) Verify, (4) Teardown 
approach described by Gerard Meszaros 
in "xUnit test patterns: Refactoring test code" (2007).

Global (module) scope setups and teardowns, however,
are conducted using pytest fixtures.
"""

# \\\\\\\\\\\\\\\\\\\\\\
#  FIXTURE DEFINITIONS
# //////////////////////

@pytest.fixture(scope="module")
def params():
    """
    Returns dict of test parameters extracted from textfile
    """
    # (1) Setup
    path = os.getenv("PARAMS_PATH")
    params = {
        "sim_dir": "",
        "si_path": "",
        "conf_path": "",
        "reviewers": ""
    }

    with open(path, "r") as f:
        for line in f.readlines():
            param_name, param_val = line.split("=") 
            params[param_name] = param_val.rstrip() # Remove newlines
    
    return params


@pytest.fixture(scope="module")
def ppt():
    """
    Returns PowerPoint singleton
    """
    # (1) Setup
    PowerPoint = win32.gencache.EnsureDispatch("PowerPoint.Application")
    yield PowerPoint

    # (4) Teardown
    PowerPoint.Quit()


@pytest.fixture(scope="module")
def conf_tools(ppt, params):
    """
    Returns ConfirmationTools singleton
    """
    # (1) Setup
    conf_tools = weaver.ConfirmationTools(ppt.Presentations.Open(params["conf_path"], WithWindow=False)) 
    yield conf_tools 

    # (4) Teardown
    conf_tools.pptx.Close()


@pytest.fixture(params=["1997-02-25", "1997-02-28", "2012-12-07"])
def mock_date(monkeypatch, request):
    """
    Monkeypatches builtins.input 
    in order to test `_get_date`
    """
    # (1) Setup
    def mock_input(prompt):
        return request.param

    monkeypatch.setattr("builtins.input", mock_input)
    yield

    # (4) Teardown
    # Revert builtins.input back to normal   
    monkeypatch.undo() 


# \\\\\\\\\\\\\\\\\\\\\\
#  FUNCTIONS
# //////////////////////

def test_get_rep_type(params):

    # (1) Setup
    expected_type = "si"

    # (2) Execution
    actual_type = weaver._get_rep_type(params["conf_path"])

    # (3) Verify
    assert actual_type == expected_type

    # (4) Teardown


def test_fetch_interfaces(params):

    # (1) Setup
    sim_dir = params["sim_dir"] 
    expected_if_list = ["RGMII"]

    # (2) Execute
    actual_if_list = weaver.fetch_interfaces(sim_dir)

    # (3) Verify
    assert isinstance(actual_if_list, list)
    assert actual_if_list == expected_if_list 

    # (4) Teardown
    

def test_load_template_paths(params):

    # (1) Setup
    expected_si_path = params["si_path"]
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


def test_get_date(mock_date):

    # (1) Setup
    expected_dates = [ "25 Feb. 1997", "28 Feb. 1997", "07 Dec. 2012" ]

    # (2) Execute
    actual_date = weaver._get_date()

    # (3) Verify
    assert actual_date in expected_dates

    # (4) Teardown


def test_init_reports(conf_tools):

    # (1) Setup
    rep_type = "si"
    interfaces = ["RGMII"]
    rep_str = f"{rep_type.upper()} Report for {interfaces[0]}"

    # (2) Execute
    reports = weaver.init_reports(rep_type, conf_tools, interfaces)
    test_rep = reports[0]

    # (3) Verify
    assert len(reports) == 1
    assert test_rep.report_type[:].lower() == rep_type
    assert test_rep.pptx # Check if Presentation obj created
    assert str(test_rep) == rep_str

    # (4) Teardown


# \\\\\\\\\\\\\\\\\\\\\\
#  CLASS METHODS
# //////////////////////

def test_get_creators(conf_tools, params):
    
    # (1) Setup
    expected_reviewers = params["reviewers"]
    
    # (2) Execute
    actual_creators = conf_tools.get_creators()

    # (3) Verify
    # Despite key names, all dict values returned
    # by conf_tools.get_creators() should be strings. 
    # This is true even when multiple persons are involved, 
    # as all names are separated by commas 
    assert isinstance(actual_creators, dict)
    assert isinstance(actual_creators["preparers"], str)
    assert actual_creators["reviewers"] == expected_reviewers
    assert not actual_creators["preparers"] == expected_reviewers
    assert actual_creators["preparers"].find(",") > -1

    # (4) Teardown


def test_get_toc(conf_tools):

    # (1) Setup
    # None

    # (2) Execute
    actual_toc = conf_tools.get_toc()

    # (3) Verify
    assert len(actual_toc.keys()) == 3
    for slide_nums in actual_toc.values():
        # All values have been found
        assert not slide_nums == "" and isinstance(slide_nums, list)
        for sn in slide_nums:
            assert isinstance(sn, int)

    # (4) Teardown


def test_get_table(conf_tools):

    # (1) Setup
    # As per the C++ MsoTriState Enum
    # Refer to https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.core.msotristate?view=office-pia 
    msoTrue = -1
    cover_slide = conf_tools.pptx.Slides(weaver.COVER_SLIDE)

    # (2) Execute
    table = conf_tools._get_table(cover_slide.Shapes)

    # (3) Verify
    # Check parent element of grabbed table
    assert not table == None
    assert table.Parent.HasTable == msoTrue

    # (4) Teardown