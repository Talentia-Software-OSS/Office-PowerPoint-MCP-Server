"""
Hyperlink management tools for PowerPoint MCP Server.
Implements hyperlink operations for text shapes and runs.
"""

from typing import Dict, List, Optional, Any
import os
from mcp.server.fastmcp import FastMCP
import utils as ppt_utils

def register_hyperlink_tools(app, resolve_presentation_path):
    """Register hyperlink management tools with the FastMCP app."""
    
    @app.tool()
    def manage_hyperlinks(
        operation: str,
        slide_index: int,
        shape_index: int = None,
        text: str = None, 
        url: str = None,
        run_index: int = 0,
        presentation_file_name: str = None
    ) -> Dict:
        """
        Manage hyperlinks in text shapes and runs.
        
        Args:
            operation: Operation type ("add", "remove", "list", "update")
            slide_index: Index of the slide (0-based)
            shape_index: Index of the shape on the slide (0-based)
            text: Text to make into hyperlink (for "add" operation)
            url: URL for the hyperlink
            run_index: Index of text run within the shape (0-based)
            presentation file name: File name or path to the presentation
            
        Returns:
            Dictionary with operation results
        """
        try:
            if not presentation_file_name:
                return {"error": "presentation_file_name is required"}
            path = resolve_presentation_path(presentation_file_name)
            if not os.path.exists(path):
                return {"error": f"File not found: {path}"}
            pres = ppt_utils.open_presentation(path)
            
            # Validate slide index
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Slide index {slide_index} out of range"}
            
            slide = pres.slides[slide_index]
            
            if operation == "list":
                # List all hyperlinks in the slide
                hyperlinks = []
                for shape_idx, shape in enumerate(slide.shapes):
                    if hasattr(shape, 'text_frame') and shape.text_frame:
                        for para_idx, paragraph in enumerate(shape.text_frame.paragraphs):
                            for run_idx, run in enumerate(paragraph.runs):
                                if run.hyperlink.address:
                                    hyperlinks.append({
                                        "shape_index": shape_idx,
                                        "paragraph_index": para_idx,
                                        "run_index": run_idx,
                                        "text": run.text,
                                        "url": run.hyperlink.address
                                    })
                
                return {
                    "message": f"Found {len(hyperlinks)} hyperlinks on slide {slide_index}",
                    "hyperlinks": hyperlinks,
                    "file_path": path
                }
            
            # For other operations, validate shape index
            if shape_index is None or not (0 <= shape_index < len(slide.shapes)):
                return {"error": f"Shape index {shape_index} out of range"}
            
            shape = slide.shapes[shape_index]
            
            # Check if shape has text
            if not hasattr(shape, 'text_frame') or not shape.text_frame:
                return {"error": "Shape does not contain text"}
            
            if operation == "add":
                if not text or not url:
                    return {"error": "Both 'text' and 'url' are required for adding hyperlinks"}
                
                # Add new text run with hyperlink
                paragraph = shape.text_frame.paragraphs[0]
                run = paragraph.add_run()
                run.text = text
                run.hyperlink.address = url
                ppt_utils.save_presentation(pres, path)
                
                return {
                    "message": f"Added hyperlink '{text}' -> '{url}' to shape {shape_index}",
                    "text": text,
                    "url": url,
                    "file_path": path
                }
            
            elif operation == "update":
                if not url:
                    return {"error": "URL is required for updating hyperlinks"}
                
                # Update existing hyperlink
                paragraphs = shape.text_frame.paragraphs
                if run_index < len(paragraphs[0].runs):
                    run = paragraphs[0].runs[run_index]
                    old_url = run.hyperlink.address
                    run.hyperlink.address = url
                    ppt_utils.save_presentation(pres, path)
                    
                    return {
                        "message": f"Updated hyperlink from '{old_url}' to '{url}'",
                        "old_url": old_url,
                        "new_url": url,
                        "text": run.text,
                        "file_path": path
                    }
                else:
                    return {"error": f"Run index {run_index} out of range"}
            
            elif operation == "remove":
                # Remove hyperlink from specific run
                paragraphs = shape.text_frame.paragraphs
                if run_index < len(paragraphs[0].runs):
                    run = paragraphs[0].runs[run_index]
                    old_url = run.hyperlink.address
                    run.hyperlink.address = None
                    ppt_utils.save_presentation(pres, path)
                    
                    return {
                        "message": f"Removed hyperlink '{old_url}' from text '{run.text}'",
                        "removed_url": old_url,
                        "text": run.text,
                        "file_path": path
                    }
                else:
                    return {"error": f"Run index {run_index} out of range"}
            
            else:
                return {"error": f"Unsupported operation: {operation}. Use 'add', 'remove', 'list', or 'update'"}
                
        except Exception as e:
            return {"error": f"Failed to manage hyperlinks: {str(e)}"}