from docx import Document
from lxml import etree
import os
import json

# Color palettes configuration
COLOR_PALETTES = {
    "red": {
        "name": "Red Palette",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "8B0000",
            "lt2": "FFE6E6",
            "accent1": "DC143C",
            "accent2": "FF0000",
            "accent3": "B22222",
            "accent4": "CD5C5C",
            "accent5": "F08080",
            "accent6": "FA8072",
            "hlink": "C71585",
            "folHlink": "8B008B"
        }
    },
    "orange": {
        "name": "Orange Palette",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "CC5500",
            "lt2": "FFE5CC",
            "accent1": "FF8C00",
            "accent2": "FF6347",
            "accent3": "FF7F50",
            "accent4": "FFA500",
            "accent5": "FFB347",
            "accent6": "FFCC99",
            "hlink": "FF4500",
            "folHlink": "D2691E"
        }
    },
    "yellow": {
        "name": "Yellow Palette",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "B8860B",
            "lt2": "FFFACD",
            "accent1": "FFD700",
            "accent2": "FFEA00",
            "accent3": "F0E68C",
            "accent4": "EEE8AA",
            "accent5": "FFEB3B",
            "accent6": "FFF59D",
            "hlink": "DAA520",
            "folHlink": "B8860B"
        }
    },
    "green": {
        "name": "Green Palette",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "006400",
            "lt2": "E6F4EA",
            "accent1": "228B22",
            "accent2": "32CD32",
            "accent3": "00A86B",
            "accent4": "66CDAA",
            "accent5": "90EE90",
            "accent6": "98FB98",
            "hlink": "008000",
            "folHlink": "2E8B57"
        }
    },
    "blue": {
        "name": "Blue Palette",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "00008B",
            "lt2": "E6F2FF",
            "accent1": "1E90FF",
            "accent2": "4169E1",
            "accent3": "0000FF",
            "accent4": "87CEEB",
            "accent5": "6495ED",
            "accent6": "ADD8E6",
            "hlink": "0066CC",
            "folHlink": "000080"
        }
    },
    "purple": {
        "name": "Purple Palette",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "4B0082",
            "lt2": "F3E5F5",
            "accent1": "8B00FF",
            "accent2": "9370DB",
            "accent3": "BA55D3",
            "accent4": "DA70D6",
            "accent5": "DDA0DD",
            "accent6": "E6B3E6",
            "hlink": "800080",
            "folHlink": "663399"
        }
    },
    "pink": {
        "name": "Pink Palette",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "C71585",
            "lt2": "FFE6F0",
            "accent1": "FF69B4",
            "accent2": "FF1493",
            "accent3": "DB7093",
            "accent4": "FFB6C1",
            "accent5": "FFC0CB",
            "accent6": "FFD9E6",
            "hlink": "FF00FF",
            "folHlink": "C71585"
        }
    },
    "grey": {
        "name": "Grey Palette",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "2F4F4F",
            "lt2": "F5F5F5",
            "accent1": "696969",
            "accent2": "808080",
            "accent3": "A9A9A9",
            "accent4": "C0C0C0",
            "accent5": "D3D3D3",
            "accent6": "DCDCDC",
            "hlink": "4F4F4F",
            "folHlink": "2F2F2F"
        }
    },
    "brown": {
        "name": "Brown Palette",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "5D4037",
            "lt2": "EFEBE9",
            "accent1": "8B4513",
            "accent2": "A0522D",
            "accent3": "CD853F",
            "accent4": "D2691E",
            "accent5": "DEB887",
            "accent6": "F5DEB3",
            "hlink": "8B4513",
            "folHlink": "654321"
        }
    },
    "ocean": {
        "name": "Ocean Breeze",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "004D5C",
            "lt2": "E0F7FA",
            "accent1": "00ACC1",
            "accent2": "26C6DA",
            "accent3": "00897B",
            "accent4": "4DD0E1",
            "accent5": "80DEEA",
            "accent6": "B2EBF2",
            "hlink": "0097A7",
            "folHlink": "006064"
        }
    },
    "sunset": {
        "name": "Sunset Vibes",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "BF360C",
            "lt2": "FFF3E0",
            "accent1": "FF6F00",
            "accent2": "FF9800",
            "accent3": "E91E63",
            "accent4": "FFAB91",
            "accent5": "FFCCBC",
            "accent6": "FFE0B2",
            "hlink": "F4511E",
            "folHlink": "BF360C"
        }
    },
    "forest": {
        "name": "Forest Path",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "1B5E20",
            "lt2": "F1F8E9",
            "accent1": "558B2F",
            "accent2": "689F38",
            "accent3": "795548",
            "accent4": "8BC34A",
            "accent5": "AED581",
            "accent6": "C5E1A5",
            "hlink": "33691E",
            "folHlink": "1B5E20"
        }
    },
    "lavender": {
        "name": "Lavender Dream",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "4A148C",
            "lt2": "F3E5F5",
            "accent1": "7B1FA2",
            "accent2": "9C27B0",
            "accent3": "AB47BC",
            "accent4": "BA68C8",
            "accent5": "CE93D8",
            "accent6": "E1BEE7",
            "hlink": "6A1B9A",
            "folHlink": "4A148C"
        }
    },
    "corporate": {
        "name": "Corporate Blue",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "1A237E",
            "lt2": "E8EAF6",
            "accent1": "283593",
            "accent2": "3F51B5",
            "accent3": "5C6BC0",
            "accent4": "7986CB",
            "accent5": "9FA8DA",
            "accent6": "C5CAE9",
            "hlink": "1976D2",
            "folHlink": "0D47A1"
        }
    },
    "mint": {
        "name": "Mint Fresh",
        "colors": {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "004D40",
            "lt2": "E0F2F1",
            "accent1": "00796B",
            "accent2": "009688",
            "accent3": "26A69A",
            "accent4": "4DB6AC",
            "accent5": "80CBC4",
            "accent6": "B2DFDB",
            "hlink": "00897B",
            "folHlink": "004D40"
        }
    }
}

def change_theme_colors(docx_path, palette_name):
    """
    Change the theme colors of a Word document to a specified palette.
    
    Args:
        docx_path: Path to the .docx file
        palette_name: Name of the color palette to apply
    """
    if palette_name not in COLOR_PALETTES:
        raise ValueError(f"Unknown palette: {palette_name}. Available: {', '.join(COLOR_PALETTES.keys())}")
    
    palette = COLOR_PALETTES[palette_name]["colors"]
    
    doc = Document(docx_path)
    
    # Access the theme part
    theme_part = doc.part.part_related_by(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
    )
    
    # Get the theme XML
    theme_xml = theme_part.blob
    theme_element = etree.fromstring(theme_xml)
    
    # Define namespaces
    namespaces = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    }
    
    # Find the color scheme element
    color_scheme = theme_element.find('.//a:clrScheme', namespaces)
    
    if color_scheme is not None:
        # Update each color in the scheme
        for color_name, rgb_value in palette.items():
            color_elem = color_scheme.find(f'.//a:{color_name}', namespaces)
            if color_elem is not None:
                # Clear existing color definitions
                color_elem.clear()
                # Set new RGB color
                srgb_clr = etree.SubElement(
                    color_elem,
                    f'{{{namespaces["a"]}}}srgbClr',
                    val=rgb_value
                )
        
        # Update the theme part with modified XML
        theme_part._blob = etree.tostring(theme_element, xml_declaration=True, encoding='UTF-8')
    
    # Save the modified document
    doc.save(docx_path)
    print(f"✓ Applied '{COLOR_PALETTES[palette_name]['name']}' theme to: {docx_path}")

def process_directory(directory='.', palette='red'):
    """
    Process all .docx files in the specified directory.
    
    Args:
        directory: Path to the directory (default: current directory)
        palette: Name of the color palette to apply (default: 'red')
    """
    processed = 0
    errors = 0
    
    print(f"Using palette: {COLOR_PALETTES[palette]['name']}")
    print(f"Processing directory: {directory}\n")
    
    # Get all .docx files in the directory
    for filename in os.listdir(directory):
        if filename.endswith('.docx') and not filename.startswith('~$'):
            filepath = os.path.join(directory, filename)
            try:
                change_theme_colors(filepath, palette)
                processed += 1
            except Exception as e:
                print(f"✗ Error processing {filename}: {str(e)}")
                errors += 1
    
    print(f"\n{'='*50}")
    print(f"Summary: {processed} documents updated, {errors} errors")
    print(f"{'='*50}")

def list_palettes():
    """List all available color palettes."""
    print("\nAvailable Color Palettes:")
    print("="*50)
    for key, palette in COLOR_PALETTES.items():
        print(f"  {key:12} - {palette['name']}")
    print("="*50)

if __name__ == "__main__":
    import sys
    
    print("Word Document Theme Color Changer")
    print("="*50)
    
    # Parse command line arguments
    if len(sys.argv) > 1:
        if sys.argv[1] == '--list':
            list_palettes()
        else:
            palette = sys.argv[1]
            directory = sys.argv[2] if len(sys.argv) > 2 else '.'
            
            if palette in COLOR_PALETTES:
                process_directory(directory, palette)
            else:
                print(f"Error: Unknown palette '{palette}'")
                list_palettes()
    else:
        print("\nUsage:")
        print("  python script.py <palette_name> [directory]")
        print("  python script.py --list")
        print("\nExample:")
        print("  python script.py ocean")
        print("  python script.py blue ./documents")
        list_palettes()