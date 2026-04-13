# -*- coding: utf-8 -*-
"""
Revit version compatibility helpers.
"""
import os

def get_revit_app_version():
    """
    Return the major Revit version number.

    Returns:
        int: For example 2025, 2024.
    """
    try:
        from Autodesk.Revit.ApplicationServices import Application
        app_version = __revit__.Application.VersionNumber
        # Format: "2025.0.12345" -> extract "2025"
        major_version = int(app_version.split('.')[0])
        return major_version
    except Exception:
        return None


def get_ifc_shared_parameters_path(revit_app_version=None):
    """
    Find the IFC shared parameters file path.

    Args:
        revit_app_version: Optional major version, for example 2025.

    Returns:
        str: The shared parameters path, or None if no file is found.
    """
    if revit_app_version is None:
        revit_app_version = get_revit_app_version()
    
    if revit_app_version is None:
        return None
    
    # Candidate paths in priority order.
    candidates = [
        # Standard path for Revit 2024+.
        r"C:\Program Files\Autodesk\Revit {}\IFC Shared Parameters-RevitIFCBuiltIn_ALL.txt".format(revit_app_version),
        # Alternative naming convention.
        r"C:\Program Files\Autodesk\Revit {}\IFC Shared Parameters.txt".format(revit_app_version),
        # ProgramData fallback.
        r"C:\ProgramData\Autodesk\Revit\Shared Parameters\IFC Shared Parameters.txt",
        # User AppData fallback.
        os.path.expanduser(r"~\AppData\Roaming\Autodesk\Revit\Shared Parameters\IFC Shared Parameters.txt"),
    ]
    
    # Try a few older versions as a fallback (down to project minimum 2024).
    for offset in range(1, 4):
        fallback_version = revit_app_version - offset
        if fallback_version >= 2024:
            candidates.append(
                r"C:\Program Files\Autodesk\Revit {}\IFC Shared Parameters-RevitIFCBuiltIn_ALL.txt".format(fallback_version)
            )
    
    for candidate_path in candidates:
        if os.path.exists(candidate_path):
            return candidate_path
    
    return None


def check_revit_version_compatible(min_version=2024):
    """
    Check whether the current Revit version is supported.

    Args:
        min_version: Minimum supported version.
    
    Returns:
        tuple: (is_compatible: bool, version_info: str)
    """
    current_version = get_revit_app_version()
    
    if current_version is None:
        return False, "Cannot detect Revit version"
    
    if current_version < min_version:
        return False, "Revit {} is too old (minimum: {})".format(current_version, min_version)
    
    return True, "Revit {} is supported".format(current_version)


def ensure_revit_version(min_version=2024):
    """
    Validate the Revit version and show a warning if needed.

    Args:
        min_version: Minimum supported version.
    
    Returns:
        bool: True if the version is compatible.
    """
    from pyrevit import forms
    
    compatible, info = check_revit_version_compatible(min_version=min_version)
    
    if not compatible:
        forms.alert(
            "External Classification Plugin - Version Check\n\n" + info,
            title="Incompatible Revit Version",
            warn_icon=True
        )
        return False
    
    return True


def get_revit_install_path():
    """
    Return the Revit installation directory.

    Returns:
        str: Path to C:\Program Files\Autodesk\Revit YYYY or None.
    """
    try:
        version = get_revit_app_version()
        if version is None:
            return None
        
        base_path = r"C:\Program Files\Autodesk\Revit {}".format(version)
        if os.path.exists(base_path):
            return base_path
        
        return None
    except Exception:
        return None


# Debugging / info
if __name__ == "__main__":
    print("=== Revit Compatibility Info ===")
    print("Revit Version: {}".format(get_revit_app_version()))
    print("IFC Shared Parameters Path: {}".format(get_ifc_shared_parameters_path()))
    print("Revit Install Path: {}".format(get_revit_install_path()))
    compatible, info = check_revit_version_compatible()
    print("Compatibility: {} - {}".format(compatible, info))
