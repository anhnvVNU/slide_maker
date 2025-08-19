#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
from pathlib import Path
import yaml

# Test loading templates and creating examples
def test_template_loading():
    template_path = Path("template")
    
    template_files = {
        "cover": "cover_slide/cover_slide.json",
        "agenda": "agenda_slide/agenda_slide.json",
        "section_divider": "section_divider/section_divider.json",
        "content": "content/content_slide.json",
        "subtitle": "subtitle/subtitle_slide.json"
    }
    
    templates = {}
    for slide_type, file_path in template_files.items():
        full_path = template_path / file_path
        print(f"Loading {slide_type} from {full_path}")
        with open(full_path, 'r', encoding='utf-8') as f:
            templates[slide_type] = json.load(f)
        print(f"  - Loaded successfully, slide_type: {templates[slide_type].get('slide_type')}")
    
    return templates

def test_representation_reading():
    represent_file = Path("output/represent.txt")
    print(f"\nReading representation from {represent_file}")
    with open(represent_file, 'r', encoding='utf-8') as f:
        content = f.read()
    print(f"  - Read {len(content)} characters")
    print(f"  - First 100 chars: {content[:100]}...")
    return content

def test_settings_loading():
    settings_path = Path("settings.yaml")
    print(f"\nLoading settings from {settings_path}")
    with open(settings_path, 'r', encoding='utf-8') as f:
        settings = yaml.safe_load(f)
    
    if 'prompts' in settings and 'slide_generator' in settings['prompts']:
        print("  - Slide generator prompts found")
        print(f"  - System prompt length: {len(settings['prompts']['slide_generator']['system'])} chars")
        print(f"  - User prompt template length: {len(settings['prompts']['slide_generator']['user'])} chars")
    else:
        print("  - WARNING: Slide generator prompts not found in settings")
    
    return settings

if __name__ == "__main__":
    print("Testing Script Generator Components")
    print("=" * 50)
    
    try:
        templates = test_template_loading()
        print(f"\nSuccessfully loaded {len(templates)} templates")
    except Exception as e:
        print(f"\nError loading templates: {e}")
    
    try:
        representation = test_representation_reading()
        print("\nSuccessfully read representation")
    except Exception as e:
        print(f"\nError reading representation: {e}")
    
    try:
        settings = test_settings_loading()
        print("\nSuccessfully loaded settings")
    except Exception as e:
        print(f"\nError loading settings: {e}")
    
    print("\n" + "=" * 50)
    print("Test complete!")