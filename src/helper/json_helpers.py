import json 

def load_bullet_points(json_file_path: str, key: str) -> list:
    """Loads the bullet points for a specific slide from a JSON file."""
    
    with open(json_file_path, 'r') as file:
        data = json.load(file)
    
    # Get the bullet points for the specified slide, return an empty list if not found
    bullet_points = data.get(key)
    
    # Format the bullet points with proper indentation
    formatted_bullet_points = format_bullet_points(bullet_points)
    
    return formatted_bullet_points

def format_bullet_points(bullet_points, indent_level=0):
    """Formats the bullet points, handling indentation for nested points."""

    formatted_points = []
    
    for point in bullet_points:
        if isinstance(point, list):
            # If the point is a nested list, recursively call format_bullet_points with increased indentation
            formatted_points.extend(format_bullet_points(point, indent_level + 1))
        else:
            # Format the bullet point with the appropriate indentation
            indentation = '     ' * indent_level  # 3 spaces per indentation level
            formatted_points.append(f"{indentation}• {point}")
    
    return formatted_points
