import os
import inspect


def float_or_none(s: str):
    try:
        return float(s)
    except ValueError:
        return None


def float_or_zero(s: str):
    try:
        return float(s)
    except ValueError:
        return 0


def int_or_none(s: str):
    try:
        return int(s)
    except ValueError:
        return None


def empty_comma_split(s: str):
    if s == "":
        return []
    return s.split(",")


def remove_prefix(text, prefix):  # present in python 3.9
    if text.startswith(prefix):
        return text[len(prefix):]
    return text


# Found on internet when win32com.client gave an error message
def dispatch(app_name: str):
    try:
        from win32com import client
        app = client.gencache.EnsureDispatch(app_name)
    except AttributeError:
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil
        # Remove cache and try again.
        module_list = [m.__name__ for m in sys.modules.values()]
        for module in module_list:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        from win32com import client
        app = client.gencache.EnsureDispatch(app_name)
    return app


def open_visum(version_filename="test.ver"):
    assert 'Visum' not in globals()
    print("opening Visum")
    result = dispatch("Visum.Visum.240")
    folder_of_this_python_file = os.path.dirname(os.path.realpath(__file__))
    folder_of_test_version_files = os.path.join(folder_of_this_python_file, r"..\TestVersionFiles")
    full_filename = os.path.join(folder_of_test_version_files, version_filename)
    print(f"opening {full_filename} in Visum")
    result.LoadVersion(full_filename)
    return result


def get_working_folder():
    return os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
