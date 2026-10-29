import os
import sys
import copy
import json
import argparse
import shutil
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml

def replace_nvidia_logo(prs, logo_path):
    NVIDIA_LOGO_SHA1 = "31f09ae9dfaa556bdf49a0d943ebbfb98743f471"
    replaced = False
    
    # 1. Search all slides for the logo image
    for slide in prs.slides:
        for rel_id, rel in slide.part.rels.items():
            if "image" in rel.reltype:
                part = rel._target
                if hasattr(part, "sha1") and part.sha1 == NVIDIA_LOGO_SHA1:
                    print(f"Replacing logo part: {part.partname} in slide relationships")
                    with open(logo_path, "rb") as f:
                        part._blob = f.read()
                    replaced = True
                    
    # 2. Search all slide layouts for the logo image
    for layout in prs.slide_layouts:
        for rel_id, rel in layout.part.rels.items():
            if "image" in rel.reltype:
                part = rel._target
                if hasattr(part, "sha1") and part.sha1 == NVIDIA_LOGO_SHA1:
                    print(f"Replacing logo part: {part.partname} in layout relationships")
                    with open(logo_path, "rb") as f:
                        part._blob = f.read()
                    replaced = True
                    
    # 3. Search all slide masters for the logo image
    for master in prs.slide_masters:
        for rel_id, rel in master.part.rels.items():
            if "image" in rel.reltype:
                part = rel._target
                if hasattr(part, "sha1") and part.sha1 == NVIDIA_LOGO_SHA1:
                    print(f"Replacing logo part: {part.partname} in master relationships")
                    with open(logo_path, "rb") as f:
                        part._blob = f.read()
                    replaced = True
                    
    return replaced

def replace_text_in_slide(slide, target, replacement):
    found = False
    for shape in slide.shapes:
        if shape.has_text_frame:
            if target in shape.text_frame.text:
                found = True
                replaced = False
                
                # Try run-level replacement first (preserves formatting of run)
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        if target in r.text:
                            r.text = r.text.replace(target, replacement)
                            replaced = True
                            
                # Fallback if target is split across multiple runs in a paragraph
                if not replaced:
                    for p in shape.text_frame.paragraphs:
                        if target in p.text:
                            # Save formatting from first run if available
                            font_name = 'Inter'
                            font_size = Pt(14)
                            bold = False
                            color = None
                            if p.runs:
                                run = p.runs[0]
                                font_name = run.font.name or font_name
                                font_size = run.font.size or font_size
                                bold = run.font.bold or bold
                                if run.font.color and run.font.color.type == 1:
                                    color = run.font.color.rgb
                            
                            # Perform full replacement
                            p.text = p.text.replace(target, replacement)
                            
                            # Re-apply styling
                            for r in p.runs:
                                r.font.name = font_name
                                r.font.size = font_size
                                r.font.bold = bold
                                if color:
                                    r.font.color.rgb = color
    return found

def clone_slide(source_slide, target_prs):
    layout = source_slide.slide_layout
    target_slide = target_prs.slides.add_slide(layout)
    
    # Remove default shapes added by default layout
    shapes_to_remove = list(target_slide.shapes)
    for shape in shapes_to_remove:
        el = shape.element
        el.getparent().remove(el)
        
    # Copy shapes from source
    for shape in source_slide.shapes:
        el = shape.element
        new_el = copy.deepcopy(el)
        target_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
        
    # Copy relationships safely using get_or_add
    rid_map = {}
    for rel_id, rel in source_slide.part.rels.items():
        if "notesSlide" not in rel.reltype:
            try:
                if rel.is_external:
                    new_rid = target_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
                else:
                    new_rid = target_slide.part.rels.get_or_add(rel.reltype, rel._target)
                rid_map[rel.rId] = new_rid
            except Exception as e:
                print(f"Rel copy error in cloning slide: {e}")
                
    # Update relationship IDs in XML string to avoid clashing/missing embeds
    xml_str = target_slide.element.xml
    for old_rid, new_rid in rid_map.items():
        if old_rid != new_rid:
            xml_str = xml_str.replace(f'r:embed="{old_rid}"', f'r:embed="{new_rid}"')
            xml_str = xml_str.replace(f'r:link="{old_rid}"', f'r:link="{new_rid}"')
            xml_str = xml_str.replace(f'r:id="{old_rid}"', f'r:id="{new_rid}"')
            
    # Apply updated XML back to target slide
    target_slide._element = parse_xml(xml_str)
    return target_slide

def find_template(template_arg, script_dir):
    repo_root = os.path.dirname(script_dir)
    candidates = [
        template_arg,
        template_arg + ".pptx" if not template_arg.endswith(".pptx") else template_arg,
        os.path.join(repo_root, "templates", "addsum", template_arg),
        os.path.join(repo_root, "templates", "addsum", f"{template_arg}.pptx"),
        os.path.join(repo_root, "templates", template_arg),
        os.path.join(repo_root, "templates", f"{template_arg}.pptx"),
        os.path.join(os.getcwd(), "templates", "addsum", template_arg),
        os.path.join(os.getcwd(), "templates", "addsum", f"{template_arg}.pptx"),
        os.path.join(os.getcwd(), template_arg),
        os.path.join(os.getcwd(), f"{template_arg}.pptx"),
        "/Users/jamesduy/Documents/Documents - James's MacBook Pro (3)/James Duy /James's Antigravity/hilab-technology/templates/addsum/" + template_arg,
        "/Users/jamesduy/Documents/Documents - James's MacBook Pro (3)/James Duy /James's Antigravity/hilab-technology/templates/addsum/" + template_arg + ".pptx",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return os.path.abspath(c)
    return None

def find_default_logo(script_dir):
    repo_root = os.path.dirname(script_dir)
    candidates = [
        os.path.join(repo_root, "resources", "logo", "hilab-white.png"),
        os.path.join(repo_root, "resources", "hilab-white.png"),
        os.path.join(os.getcwd(), "resources", "logo", "hilab-white.png"),
        os.path.join(os.getcwd(), "hilab-white.png"),
        "/Users/jamesduy/Documents/Documents - James's MacBook Pro (3)/James Duy /James's Antigravity/hilab-technology/hilab-white.png",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return os.path.abspath(c)
    return None

def main():
    parser = argparse.ArgumentParser(description="Multiple Slide Builder CLI")
    parser.add_argument("--template", required=True, help="Template name (e.g. company-intro) or direct path to .pptx")
    parser.add_argument("--config", required=True, help="Path to JSON configuration file containing slide definition and replacements")
    parser.add_argument("--output", required=True, help="Path where output .pptx will be saved")
    parser.add_argument("--logo", help="Path to logo PNG file to replace the default green NVIDIA logo image")
    
    args = parser.parse_args()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 1. Resolve template path
    template_path = find_template(args.template, script_dir)
    if not template_path:
        print(f"Error: Template not found for '{args.template}'. Checked relative and default template paths.")
        sys.exit(1)
            
    print(f"Loading template: {template_path}")
    
    # 2. Parse config file
    print(f"Loading configuration: {args.config}")
    with open(args.config, 'r', encoding='utf-8') as f:
        config = json.load(f)
        
    slides_config = config.get("slides", [])
    if not slides_config:
        print("Warning: No slides defined in configuration.")
        
    # 3. Create target presentation
    output_dir = os.path.dirname(os.path.abspath(args.output))
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        
    shutil.copy(template_path, args.output)
    prs = Presentation(args.output)
    original_slides_count = len(prs.slides)
    
    # 4. Clone slides dynamically as specified in config (allowing duplication and reordering)
    new_slides = []
    for idx, slide_def in enumerate(slides_config):
        src_index = slide_def.get("source_index")
        if src_index is None or src_index < 0 or src_index >= original_slides_count:
            print(f"Error: Invalid source_index {src_index} in slide config {idx}.")
            continue
            
        source_slide = prs.slides[src_index]
        print(f"Duplicating/Cloning slide layout index {src_index} for slide {idx + 1}...")
        new_slide = clone_slide(source_slide, prs)
        
        # Apply text replacements on the newly cloned slide
        replacements = slide_def.get("replacements", {})
        for target, replacement in replacements.items():
            replace_text_in_slide(new_slide, target, replacement)
            
        new_slides.append(new_slide)
        
    # 5. Delete the original template slides safely to prevent PowerPoint corruption
    print(f"Deleting the {original_slides_count} original template slides safely...")
    for i in range(original_slides_count - 1, -1, -1):
        slide = prs.slides[i]
        rId = None
        for rel_id, rel in prs.part.rels.items():
            if rel.target_part == slide.part:
                rId = rel_id
                break
        if rId:
            prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[i]
        
    # 6. Replace Nvidia Logo if logo path is provided/resolved
    logo_file = args.logo
    if not logo_file:
        logo_file = find_default_logo(script_dir)
            
    if logo_file and os.path.isfile(logo_file):
        print(f"Replacing template logo with logo from: {logo_file}")
        replace_nvidia_logo(prs, logo_file)
    else:
        print("Note: No replacement logo applied (logo file not found or not specified).")
        
    # 7. Save output presentation
    prs.save(args.output)
    print(f"Successfully generated slides at: {args.output}")

if __name__ == "__main__":
    main()
