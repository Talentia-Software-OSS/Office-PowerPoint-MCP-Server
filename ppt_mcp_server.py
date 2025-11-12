#!/usr/bin/env python
"""
MCP Server for PowerPoint manipulation using python-pptx.
Consolidated version with 20 tools organized into multiple modules.
"""
import os
import argparse
from typing import Dict, Any
from mcp.server.fastmcp import FastMCP

# import utils  # Currently unused
import platform
from tools import (
    register_presentation_tools,
    register_content_tools,
    register_structural_tools,
    register_professional_tools,
    # register_template_tools,
    register_hyperlink_tools,
    register_chart_tools,
    register_connector_tools,
    register_master_tools,
    register_transition_tools
)

# Initialize the FastMCP server
app = FastMCP(
    name="ppt-mcp-server"
)

# ---- Presentations Root Configuration (stateless file resolver) ----
_presentations_root_cli = None  # set via CLI parsing

def get_presentations_root() -> str:
    """
    Resolve the base directory for storing/loading presentation files.
    Priority: CLI flag -> env var PPT_PRESENTATIONS_ROOT -> ./presentations
    """
    if _presentations_root_cli:
        return _presentations_root_cli
    return os.environ.get("PPT_PRESENTATIONS_ROOT", "./presentations")

def resolve_presentation_path(name_or_path: str) -> str:
    """
    Resolve a user-supplied file name or path to an absolute path.
    - If input includes a directory or is absolute, use as-is (expanded/absolute).
    - If only a file name, place it under the presentations root directory.
    Ensures the root directory exists.
    """
    try:
        raw = name_or_path.strip()
        if os.path.isabs(raw) or os.path.dirname(raw):
            abs_path = os.path.abspath(raw)
            # Ensure directory exists for absolute or directory-containing paths
            base_name = os.path.basename(abs_path)
            looks_like_file = '.' in base_name and not base_name.startswith('.')
            dir_to_create = os.path.dirname(abs_path) if looks_like_file else abs_path
            if dir_to_create:
                os.makedirs(dir_to_create, exist_ok=True)
            return abs_path
        root = os.path.abspath(get_presentations_root())
        os.makedirs(root, exist_ok=True)
        return os.path.join(root, raw)
    except Exception as e:
        print(f"Error resolving presentation path: {e}")
        raise

# Global state to store presentations in memory
# presentations = {}
# current_presentation_id = None

# Template configuration
def get_template_search_directories():
    """
    Get list of directories to search for templates.
    Uses environment variable PPT_TEMPLATE_PATH if set, otherwise uses default directories.
    
    Returns:
        List of directories to search for templates
    """
    template_env_path = os.environ.get('PPT_TEMPLATE_PATH')
    
    if template_env_path:
        # If environment variable is set, use it as the primary template directory
        # Support multiple paths separated by colon (Unix) or semicolon (Windows)
        separator = ';' if platform.system() == "Windows" else ':'
        env_dirs = [path.strip() for path in template_env_path.split(separator) if path.strip()]
        
        # Verify that the directories exist
        valid_env_dirs = []
        for dir_path in env_dirs:
            expanded_path = os.path.expanduser(dir_path)
            if os.path.exists(expanded_path) and os.path.isdir(expanded_path):
                valid_env_dirs.append(expanded_path)
        
        if valid_env_dirs:
            # Add default fallback directories
            return valid_env_dirs + ['.', './templates', './assets', './resources']
        else:
            print(f"Warning: PPT_TEMPLATE_PATH directories not found: {template_env_path}")
    
    # Default search directories when no environment variable or invalid paths
    return ['.', './templates', './assets', './resources']

# ---- Helper Functions ----

# def get_current_presentation():
#     """Get the current presentation object or raise an error if none is loaded."""
#     if current_presentation_id is None or current_presentation_id not in presentations:
#         raise ValueError("No presentation is currently loaded. Please create or open a presentation first.")
#     return presentations[current_presentation_id]

# def get_current_presentation_id():
#     """Get the current presentation ID."""
#     return current_presentation_id

# def set_current_presentation_id(pres_id):
#     """Set the current presentation ID."""
#     global current_presentation_id
#     current_presentation_id = pres_id

def validate_parameters(params):
    """
    Validate parameters against constraints.
    
    Args:
        params: Dictionary of parameter name: (value, constraints) pairs
        
    Returns:
        (True, None) if all valid, or (False, error_message) if invalid
    """
    for param_name, (value, constraints) in params.items():
        for constraint_func, error_msg in constraints:
            if not constraint_func(value):
                return False, f"Parameter '{param_name}': {error_msg}"
    return True, None

def is_positive(value):
    """Check if a value is positive."""
    return value > 0

def is_non_negative(value):
    """Check if a value is non-negative."""
    return value >= 0

def is_in_range(min_val, max_val):
    """Create a function that checks if a value is in a range."""
    return lambda x: min_val <= x <= max_val

def is_in_list(valid_list):
    """Create a function that checks if a value is in a list."""
    return lambda x: x in valid_list

def is_valid_rgb(color_list):
    """Check if a color list is a valid RGB tuple."""
    if not isinstance(color_list, list) or len(color_list) != 3:
        return False
    return all(isinstance(c, int) and 0 <= c <= 255 for c in color_list)

def add_shape_direct(slide, shape_type: str, left: float, top: float, width: float, height: float) -> Any:
    """
    Add an auto shape to a slide using direct integer values instead of enum objects.
    
    This implementation provides a reliable alternative that bypasses potential 
    enum-related issues in the python-pptx library.
    
    Args:
        slide: The slide object
        shape_type: Shape type string (e.g., 'rectangle', 'oval', 'triangle')
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        
    Returns:
        The created shape
    """
    from pptx.util import Inches
    
    # Direct mapping of shape types to their integer values
    # These values are directly from the MS Office VBA documentation
    shape_type_map = {
        'rectangle': 1,
        'rounded_rectangle': 2, 
        'oval': 9,
        'diamond': 4,
        'triangle': 5,  # This is ISOSCELES_TRIANGLE
        'right_triangle': 6,
        'pentagon': 56,
        'hexagon': 10,
        'heptagon': 11,
        'octagon': 12,
        'star': 12,  # This is STAR_5_POINTS (value 12)
        'arrow': 13,
        'cloud': 35,
        'heart': 21,
        'lightning_bolt': 22,
        'sun': 23,
        'moon': 24,
        'smiley_face': 17,
        'no_symbol': 19,
        'flowchart_process': 112,
        'flowchart_decision': 114,
        'flowchart_data': 115,
        'flowchart_document': 119
    }
    
    # Check if shape type is valid before trying to use it
    shape_type_lower = str(shape_type).lower()
    if shape_type_lower not in shape_type_map:
        available_shapes = ', '.join(sorted(shape_type_map.keys()))
        raise ValueError(f"Unsupported shape type: '{shape_type}'. Available shape types: {available_shapes}")
    
    # Get the integer value for the shape type
    shape_value = shape_type_map[shape_type_lower]
    
    # Create the shape using the direct integer value
    try:
        # The integer value is passed directly to add_shape
        shape = slide.shapes.add_shape(
            shape_value, Inches(left), Inches(top), Inches(width), Inches(height)
        )
        return shape
    except Exception as e:
        raise ValueError(f"Failed to create '{shape_type}' shape using direct value {shape_value}: {str(e)}")

# ---- Register Tools ----
register_presentation_tools(
    app, 
    get_template_search_directories,
    resolve_presentation_path
)

register_content_tools(
    app,
    validate_parameters,
    is_positive,
    is_non_negative,
    is_in_range,
    is_valid_rgb,
    resolve_presentation_path
)

register_structural_tools(
    app,
    validate_parameters,
    is_positive,
    is_non_negative,
    is_in_range,
    is_valid_rgb,
    add_shape_direct,
    resolve_presentation_path
)

register_professional_tools(
    app,
    resolve_presentation_path
)

# register_template_tools(
#     app,
#     resolve_presentation_path
# )

register_hyperlink_tools(
    app,
    resolve_presentation_path
)

register_chart_tools(
    app,
    resolve_presentation_path
)


register_connector_tools(
    app,
    resolve_presentation_path
)

register_master_tools(
    app,
    resolve_presentation_path
)

register_transition_tools(
    app,
    resolve_presentation_path
)


# ---- Additional Utility Tools ----

@app.tool()
def get_server_info() -> Dict:
    """Get information about the MCP server."""
    return {
        "name": "PowerPoint MCP Server - Enhanced Edition",
        "version": "2.1.0",
        "total_tools": 30,  # Updated after removing state-only tools
        "features": [
            "Presentation Management (7 tools)",
            "Content Management (6 tools)", 
            "Template Operations (7 tools)",
            "Structural Elements (4 tools)",
            "Professional Design (3 tools)",
            "Specialized Features (5 tools)"
        ],
        "improvements": [
            "32 specialized tools organized into 11 focused modules",
            "68+ utility functions across 7 organized utility modules",
            "Enhanced parameter handling and validation",
            "Unified operation interfaces with comprehensive coverage",
            "Advanced template system with auto-generation capabilities",
            "Professional design tools with multiple effects and styling",
            "Specialized features including hyperlinks, connectors, slide masters",
            "Dynamic text sizing and intelligent wrapping",
            "Advanced visual effects and styling",
            "Content-aware optimization and validation",
            "Complete PowerPoint lifecycle management",
            "Modular architecture for better maintainability"
        ],
        "new_enhanced_features": [
            "Hyperlink Management - Add, update, remove, and list hyperlinks in text",
            "Advanced Chart Data Updates - Replace chart data with new categories and series",
            "Advanced Text Run Formatting - Apply formatting to specific text runs",
            "Shape Connectors - Add connector lines and arrows between points",
            "Slide Master Management - Access and manage slide masters and layouts",
            "Slide Transitions - Basic transition management (placeholder for future)"
        ]
    }

# ---- Main Function ----
def main(transport: str = "stdio", port: int = 8000):
    if transport == "http":
        import asyncio
        # Set the port for HTTP transport
        app.settings.port = port
        # Start the FastMCP server with HTTP transport
        try:
            app.run(transport='streamable-http')
        except asyncio.exceptions.CancelledError:
            print("Server stopped by user.")
        except KeyboardInterrupt:
            print("Server stopped by user.")
        except Exception as e:
            print(f"Error starting server: {e}")
            
    elif transport == "sse":
        # Run the FastMCP server in SSE (Server Side Events) mode
        app.run(transport='sse')
        
    else:
        # Run the FastMCP server
        app.run(transport='stdio')

if __name__ == "__main__":
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="MCP Server for PowerPoint manipulation using python-pptx")

    parser.add_argument(
        "-t",
        "--transport",
        type=str,
        default="stdio",
        choices=["stdio", "http", "sse"],
        help="Transport method for the MCP server (default: stdio)"
    )

    parser.add_argument(
        "-p",
        "--port",
        type=int,
        default=8000,
        help="Port to run the MCP server on (default: 8000)"
    )

    parser.add_argument(
        "-r",
        "--presentations-root",
        type=str,
        default=None,
        help="Root directory for storing/loading presentations (default: ./presentations or PPT_PRESENTATIONS_ROOT)"
    )
    args = parser.parse_args()
    # Set CLI-provided presentations root if specified
    if args.presentations_root:
        _presentations_root_cli = args.presentations_root
    main(args.transport, args.port)
