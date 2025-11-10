"""
Presentation management tools for PowerPoint MCP Server.
Handles presentation creation, opening, saving, and core properties.
"""
from typing import Dict, List, Optional, Any
import os
from mcp.server.fastmcp import FastMCP
import utils as ppt_utils


def register_presentation_tools(app: FastMCP, get_template_search_directories, resolve_presentation_path):
    """Register presentation management tools with the FastMCP app"""
    
    @app.tool()
    def create_presentation(presentation_file_name: str) -> Dict:
        """Create a new PowerPoint presentation and save to disk."""
        path = resolve_presentation_path(presentation_file_name)
        pres = ppt_utils.create_presentation()
        saved_path = ppt_utils.save_presentation(pres, path)
        return {
            "message": f"Created new presentation at: {saved_path}",
            "file_path": saved_path,
            "slide_count": len(pres.slides),
        }

    @app.tool()
    def create_presentation_from_template(template_path: str, presentation_file_name: str) -> Dict:
        """Create a new PowerPoint presentation from a template file and save to disk."""
        # Check if template file exists
        if not os.path.exists(template_path):
            # Try to find the template by searching in configured directories
            search_dirs = get_template_search_directories()
            template_name = os.path.basename(template_path)
            
            for directory in search_dirs:
                potential_path = os.path.join(directory, template_name)
                if os.path.exists(potential_path):
                    template_path = potential_path
                    break
            else:
                env_path_info = f" (PPT_TEMPLATE_PATH: {os.environ.get('PPT_TEMPLATE_PATH', 'not set')})" if os.environ.get('PPT_TEMPLATE_PATH') else ""
                return {
                    "error": f"Template file not found: {template_path}. Searched in {', '.join(search_dirs)}{env_path_info}"
                }
        
        # Create presentation from template
        try:
            pres = ppt_utils.create_presentation_from_template(template_path)
        except Exception as e:
            return {
                "error": f"Failed to create presentation from template: {str(e)}"
            }

        # Save to resolved path
        path = resolve_presentation_path(presentation_file_name)
        saved_path = ppt_utils.save_presentation(pres, path)
        return {
            "message": f"Created new presentation from template '{template_path}'",
            "template_path": template_path,
            "file_path": saved_path,
            "slide_count": len(pres.slides),
            "layout_count": len(pres.slide_layouts)
        }

    @app.tool()
    def open_presentation(presentation_file_name: str) -> Dict:
        """Open an existing presentation (stateless) and return basic info."""
        path = resolve_presentation_path(presentation_file_name)
        if not os.path.exists(path):
            return {"error": f"File not found: {path}"}
        try:
            pres = ppt_utils.open_presentation(path)
        except Exception as e:
            return {"error": f"Failed to open presentation: {str(e)}"}
        return {
            "message": f"Opened presentation: {path}",
            "file_path": path,
            "slide_count": len(pres.slides),
        }

    @app.tool()
    def save_presentation(presentation_file_name: str, file_path: Optional[str] = None) -> Dict:
        """Re-save a presentation to disk, optionally to a new path (stateless)."""
        src_path = resolve_presentation_path(presentation_file_name)
        if not os.path.exists(src_path):
            return {"error": f"File not found: {src_path}"}
        try:
            pres = ppt_utils.open_presentation(src_path)
            dest_path = os.path.abspath(file_path) if file_path else src_path
            saved_path = ppt_utils.save_presentation(pres, dest_path)
            return {"message": f"Presentation saved to {saved_path}", "file_path": saved_path}
        except Exception as e:
            return {"error": f"Failed to save presentation: {str(e)}"}

    @app.tool()
    def get_presentation_info(presentation_file_name: str) -> Dict:
        """Get information about a presentation file (stateless)."""
        path = resolve_presentation_path(presentation_file_name)
        if not os.path.exists(path):
            return {"error": f"File not found: {path}"}
        try:
            pres = ppt_utils.open_presentation(path)
            info = ppt_utils.get_presentation_info(pres)
            info["file_path"] = path
            return info
        except Exception as e:
            return {"error": f"Failed to get presentation info: {str(e)}"}

    @app.tool()
    def get_template_file_info(template_path: str) -> Dict:
        """Get information about a template file including layouts and properties."""
        # Check if template file exists
        if not os.path.exists(template_path):
            # Try to find the template by searching in configured directories
            search_dirs = get_template_search_directories()
            template_name = os.path.basename(template_path)
            
            for directory in search_dirs:
                potential_path = os.path.join(directory, template_name)
                if os.path.exists(potential_path):
                    template_path = potential_path
                    break
            else:
                return {
                    "error": f"Template file not found: {template_path}. Searched in {', '.join(search_dirs)}"
                }
        
        try:
            return ppt_utils.get_template_info(template_path)
        except Exception as e:
            return {
                "error": f"Failed to get template info: {str(e)}"
            }

    @app.tool()
    def set_core_properties(
        title: Optional[str] = None,
        subject: Optional[str] = None,
        author: Optional[str] = None,
        keywords: Optional[str] = None,
        comments: Optional[str] = None,
        presentation_file_name: Optional[str] = None
    ) -> Dict:
        """Set core document properties."""
        if not presentation_file_name:
            return {"error": "presentation_file_name is required"}
        path = resolve_presentation_path(presentation_file_name)
        if not os.path.exists(path):
            return {"error": f"File not found: {path}"}
        pres = ppt_utils.open_presentation(path)
        try:
            ppt_utils.set_core_properties(
                pres,
                title=title,
                subject=subject,
                author=author,
                keywords=keywords,
                comments=comments
            )
            ppt_utils.save_presentation(pres, path)
            return {
                "message": "Core properties updated successfully",
                "file_path": path
            }
        except Exception as e:
            return {
                "error": f"Failed to set core properties: {str(e)}"
            }