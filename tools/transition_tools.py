"""
Slide transition management tools for PowerPoint MCP Server.
Implements slide transition and timing capabilities.
"""

from typing import Dict, List, Optional, Any
import os
from mcp.server.fastmcp import FastMCP
import utils as ppt_utils

def register_transition_tools(app, resolve_presentation_path):
    """Register slide transition management tools with the FastMCP app."""
    
    @app.tool()
    def manage_slide_transitions(
        slide_index: int,
        operation: str,
        transition_type: str = None,
        duration: float = 1.0,
        presentation_file_name: str = None
    ) -> Dict:
        """
        Manage slide transitions and timing.
        
        Args:
            slide_index: Index of the slide (0-based)
            operation: Operation type ("set", "remove", "get")
            transition_type: Type of transition (basic support)
            duration: Duration of transition in seconds
            presentation file name: File name or path to the presentation
            
        Returns:
            Dictionary with transition information
        """
        try:
            if not presentation_file_name:
                return {"error": "presentation_file_name is required"}
            path = resolve_presentation_path(presentation_file_name)
            if not os.path.exists(path):
                return {"error": f"File not found: {presentation_file_name}"}
            pres = ppt_utils.open_presentation(path)
            
            # Validate slide index
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Slide index {slide_index} out of range"}
            
            slide = pres.slides[slide_index]
            
            if operation == "get":
                # Get current transition info (limited python-pptx support)
                return {
                    "message": f"Transition info for slide {slide_index}",
                    "slide_index": slide_index,
                    "note": "Transition reading has limited support in python-pptx"
                }
            
            elif operation == "set":
                result = {
                    "message": f"Transition setting requested for slide {slide_index}",
                    "slide_index": slide_index,
                    "transition_type": transition_type,
                    "duration": duration,
                    "note": "Transition setting has limited support in python-pptx - this is a placeholder for future enhancement"
                }
                # Placeholder: save to persist any possible metadata changes in future support
                ppt_utils.save_presentation(pres, path)
                return result
            
            elif operation == "remove":
                result = {
                    "message": f"Transition removal requested for slide {slide_index}",
                    "slide_index": slide_index,
                    "note": "Transition removal has limited support in python-pptx - this is a placeholder for future enhancement"
                }
                ppt_utils.save_presentation(pres, path)
                return result
            
            else:
                return {"error": f"Unsupported operation: {operation}. Use 'set', 'remove', or 'get'"}
                
        except Exception as e:
            return {"error": f"Failed to manage slide transitions: {str(e)}"}