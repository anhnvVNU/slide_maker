#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
from typing import List, Dict, Any
from pathlib import Path
import yaml
from openai import OpenAI
from dotenv import load_dotenv

class SlideScriptGenerator:
    def __init__(self, settings_path: str = "settings.yaml"):
        # Load environment variables from .env file
        load_dotenv(Path(__file__).parent.parent / '.env')
        
        self.settings_path = Path(__file__).parent.parent / settings_path
        with open(self.settings_path, 'r', encoding='utf-8') as f:
            self.settings = yaml.safe_load(f)
        
        self.client = OpenAI(
            api_key=os.getenv("OPENAI_API_KEY")
        )
        
        self.template_path = Path(__file__).parent.parent / "template"
        self.output_path = Path(__file__).parent.parent / "output"
        self.slide_templates = self._load_templates()
    
    def _load_templates(self) -> Dict[str, Any]:
        """Load all slide templates from JSON files"""
        templates = {}
        template_files = {
            "cover": "cover_slide/cover_slide.json",
            "agenda": "agenda_slide/agenda_slide.json",
            "section_divider": "section_divider/section_divider.json",
            "content": "content/content_slide.json",
            "subtitle": "subtitle/subtitle_slide.json"
        }
        
        for slide_type, file_path in template_files.items():
            full_path = self.template_path / file_path
            with open(full_path, 'r', encoding='utf-8') as f:
                templates[slide_type] = json.load(f)
        
        return templates
    
    def generate_from_structured_format(self, structured_path: str, output_path: str) -> List[Dict]:
        """Generate slides từ structured_format.txt - FUNCTION CALLING approach"""
        print("🔧 Using FUNCTION CALLING to paste content into templates...")
        
        # Read structured format
        with open(structured_path, 'r', encoding='utf-8') as f:
            structured_content = f.read()
        
        # Parse content để extract text cho từng slide
        content_data = self._parse_structured_content(structured_content)
        
        # Generate slides bằng cách paste text vào templates có sẵn
        slides = self._paste_content_to_templates(content_data)
        
        # Save to JSON
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(slides, f, ensure_ascii=False, indent=2)
        
        print(f"✓ Generated {len(slides)} slides using FUNCTION CALLING")
        return slides
    
    def _detect_content_sections(self, lines: List[str]) -> Dict[str, int]:
        """Dynamically detect content sections and task count"""
        sections = {"tasks": 0, "sections": []}
        
        for line in lines:
            line = line.strip()
            # Auto-detect task count
            if "Task" in line and "Subtitle" in line:
                sections["tasks"] += 1
            # Auto-detect section numbers 
            elif "Section Divider" in line:
                if "01" in line: sections["sections"].append("purpose")
                elif "02" in line: sections["sections"].append("completed") 
                elif "03" in line: sections["sections"].append("future")
        
        return sections
    
    def _parse_structured_content(self, structured_content: str) -> Dict:
        """Parse structured format để extract nội dung CHI TIẾT cho từng slide - DYNAMIC APPROACH"""
        lines = structured_content.split('\n')
        
        # DYNAMIC: Auto-detect sections and task count
        section_info = self._detect_content_sections(lines)
        print(f"🔍 Auto-detected: {section_info['tasks']} tasks, sections: {section_info['sections']}")
        
        content = {
            "date": "",
            "task_list": [],
            "tasks": [],
            "future_work": "",
            "detected_info": section_info  # Store dynamic info
        }
        
        current_task = None
        current_content_type = None
        current_content_buffer = []
        
        for i, line in enumerate(lines):
            line = line.strip()
            
            if "Report Date:" in line:
                content["date"] = line.split(":")[1].strip()
            
            elif "Purpose Overview" in line or "タスクの目的" in line:
                # Start collecting task list for Purpose slide  
                current_content_type = "purpose_tasks"
                current_content_buffer = []
            
            elif "●" in line and current_content_type == "purpose_tasks":
                # Task list ở phần Purpose
                content["task_list"].append(line)
            
            elif "○" in line and current_content_type == "purpose_tasks":
                # Sub-points for last task in task_list
                if content["task_list"]:
                    content["task_list"].append(line)
            
            elif "Task" in line and "Subtitle" in line:
                # Save previous task nếu có
                if current_task and current_content_buffer:
                    self._save_current_content(current_task, current_content_type, current_content_buffer)
                if current_task:
                    content["tasks"].append(current_task)
                
                # Start new task
                current_task = {
                    "title": "",
                    "problem": "",
                    "solution": "",
                    "result": ""
                }
                current_content_type = "subtitle"
                current_content_buffer = []
                
            elif current_task and line.startswith("- Content:") and current_content_type == "subtitle":
                # Extract task title từ subtitle content
                title_part = line.replace("- Content:", "").strip()
                if title_part.startswith('"') and title_part.endswith('"'):
                    title_part = title_part[1:-1]  # Remove quotes
                current_task["title"] = title_part
                
            elif "Problem Description" in line:
                # Save previous content
                if current_task and current_content_buffer:
                    self._save_current_content(current_task, current_content_type, current_content_buffer)
                current_content_type = "problem"
                current_content_buffer = []
                
            elif "Solution Process" in line:
                # Save previous content
                if current_task and current_content_buffer:
                    self._save_current_content(current_task, current_content_type, current_content_buffer)
                current_content_type = "solution"
                current_content_buffer = []
                
            elif "Results & Analysis" in line:
                # Save previous content
                if current_task and current_content_buffer:
                    self._save_current_content(current_task, current_content_type, current_content_buffer)
                current_content_type = "result"
                current_content_buffer = []
                
            elif ("●" in line or "○" in line) and current_content_type in ["problem", "solution", "result"]:
                # Collect bullet content
                current_content_buffer.append(line)
            
            elif "Future Work" in line or "次の作業内容" in line:
                current_content_type = "future_work" 
                current_content_buffer = []
                
            elif "●" in line and current_content_type == "future_work":
                content["future_work"] = line
        
        # Save last task
        if current_task and current_content_buffer:
            self._save_current_content(current_task, current_content_type, current_content_buffer)
        if current_task:
            content["tasks"].append(current_task)
        
        detected_tasks = content['detected_info']['tasks'] 
        actual_tasks = len(content['tasks'])
        print(f"🔍 Parsed {actual_tasks} tasks (detected: {detected_tasks}), {len(content['task_list'])} task list items")
        return content
    
    def _save_current_content(self, task: Dict, content_type: str, content_buffer: List[str]):
        """Save current content buffer vào task"""
        if not content_buffer:
            return
            
        content_text = "\n".join(content_buffer)
        
        if content_type == "problem":
            task["problem"] = content_text
        elif content_type == "solution":
            task["solution"] = content_text
        elif content_type == "result":
            task["result"] = content_text
    
    def _paste_content_to_templates(self, content_data: Dict) -> List[Dict]:
        """Paste text content vào JSON templates có sẵn - DYNAMIC FUNCTION CALLING"""
        slides = []
        
        # DYNAMIC: Get detected info
        detected_info = content_data.get('detected_info', {'tasks': len(content_data['tasks']), 'sections': ['purpose', 'completed', 'future']})
        total_tasks = len(content_data['tasks'])
        print(f"🚀 Generating slides for {total_tasks} tasks dynamically...")
        
        # 1. Cover slide
        cover_slide = json.loads(json.dumps(self.slide_templates["cover"]))
        self._paste_cover_slide(cover_slide, content_data['date'])
        slides.append(cover_slide)
        
        # 2. Agenda slide
        agenda_slide = json.loads(json.dumps(self.slide_templates["agenda"]))
        slides.append(agenda_slide)
        
        # 3. Section 01
        sec01 = json.loads(json.dumps(self.slide_templates["section_divider"]))
        self._paste_section_divider(sec01, "01", "タスクの目的")
        slides.append(sec01)
        
        # 4. Purpose content với task list + sub-points
        purpose_slide = json.loads(json.dumps(self.slide_templates["content"]))
        task_list_text = "\n".join(content_data["task_list"])
        self._paste_text_to_slide(purpose_slide, "1. タスクの目的", task_list_text, slide_type="content")
        slides.append(purpose_slide)
        
        # 5. Section 02
        sec02 = json.loads(json.dumps(self.slide_templates["section_divider"]))
        self._paste_section_divider(sec02, "02", "完了済み作業")
        slides.append(sec02)
        
        # 6-n. Task slides
        for i, task in enumerate(content_data["tasks"], 1):
            # Subtitle
            subtitle = json.loads(json.dumps(self.slide_templates["subtitle"]))
            self._paste_text_to_slide(subtitle, task.get('title', ''), slide_type="subtitle")
            slides.append(subtitle)
            
            # Problem slide
            if task.get("problem"):
                problem_slide = json.loads(json.dumps(self.slide_templates["content"]))
                self._paste_text_to_slide(problem_slide, "2. 完了済み作業", task["problem"], slide_type="content")
                slides.append(problem_slide)
            
            # Solution slide(s) - with balanced content distribution
            if task.get("solution"):
                solution_chunks = self._balance_content_distribution(task["solution"], max_sub_points=3)
                for chunk in solution_chunks:
                    solution_slide = json.loads(json.dumps(self.slide_templates["content"]))
                    self._paste_text_to_slide(solution_slide, "2. 完了済み作業", chunk, slide_type="content")
                    slides.append(solution_slide)
            
            # Result slide
            if task.get("result"):
                result_slide = json.loads(json.dumps(self.slide_templates["content"]))
                self._paste_text_to_slide(result_slide, "2. 完了済み作業", task["result"], slide_type="content")
                slides.append(result_slide)
        
        # DYNAMIC: Final slides based on detected sections
        if "future" in detected_info['sections']:
            sec03 = json.loads(json.dumps(self.slide_templates["section_divider"]))
            self._paste_section_divider(sec03, "03", "次の作業内容")
            slides.append(sec03)
            
            if content_data.get("future_work"):
                future_slide = json.loads(json.dumps(self.slide_templates["content"]))
                self._paste_text_to_slide(future_slide, "3. 次の作業内容", content_data["future_work"], slide_type="content")
                slides.append(future_slide)
        
        print(f"✨ Generated {len(slides)} slides dynamically (was expecting ~{total_tasks * 3 + 6} slides)")
        
        return slides
    
    def _paste_text_to_slide(self, slide: Dict, title: str, content: str = None, slide_type: str = "content"):
        """Function calling để paste text vào ĐÚNG VỊ TRÍ"""
        
        if slide_type == "content":
            # Content slide structure
            shapes = slide.get("shapes", [])
            
            for shape in shapes:
                pos_x = shape.get("position", {}).get("x")
                
                # Title shape (position x: 584875)
                if shape.get("type") == "text" and pos_x == 584875:
                    shape["paragraphs"][0]["runs"][0]["text"] = title
                
                # Header box (position x: 2653144 or has fill)
                elif shape.get("type") == "text" and (pos_x == 2653144 or shape.get("fill")):
                    shape["paragraphs"][0]["runs"][0]["text"] = title
                
                # Content bullets (position x: 626250)
                elif content and shape.get("type") == "text" and pos_x == 626250:
                    paragraphs = self._parse_bullet_content(content)
                    shape["paragraphs"] = paragraphs
        
        elif slide_type == "subtitle":
            # Subtitle slide - có 2 paragraphs
            shapes = slide.get("shapes", [])
            if len(shapes) >= 1 and title:
                shape = shapes[0]
                if shape.get("paragraphs"):
                    # Parse title to split number/code and description
                    if ". " in title and title.startswith("2."):
                        parts = title.split(". ", 1)
                        number_part = parts[0] + ". "
                        desc_part = parts[1] if len(parts) > 1 else ""
                        
                        # Split description further if it has code
                        if "】" in desc_part:
                            code_end = desc_part.find("】") + 1
                            code_part = desc_part[:code_end]
                            desc_only = desc_part[code_end:].lstrip(": ")
                            
                            # Update paragraph 1: number + code
                            if len(shape["paragraphs"]) >= 1:
                                shape["paragraphs"][0]["runs"][0]["text"] = number_part + code_part
                            
                            # Update paragraph 2: description only
                            if len(shape["paragraphs"]) >= 2:
                                shape["paragraphs"][1]["runs"][0]["text"] = desc_only.strip()
                        else:
                            shape["paragraphs"][0]["runs"][0]["text"] = title
                    else:
                        shape["paragraphs"][0]["runs"][0]["text"] = title
    
    def _parse_bullet_content(self, content: str) -> List[Dict]:
        """Parse bullet content thành paragraphs với level đúng"""
        lines = content.split('\n')
        paragraphs = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            level = 0
            bullet_char = "●"
            text = line
            
            if line.startswith("● "):
                level = 0
                bullet_char = "●"
                text = line[2:].strip()
            elif "○" in line:
                level = 1
                bullet_char = "○"
                # Extract text after ○ symbol, remove leading spaces
                text = line.split("○", 1)[-1].strip()
            
            paragraph = {
                "runs": [{
                    "text": text,
                    "font": {
                        "name": "Noto Sans JP",
                        "size_pt": 14.0
                    }
                }],
                "alignment": "left",
                "level": level,
                "line_spacing": {"value": 1.15, "type": "points"},
                "bullet": {"type": "bullet", "char": bullet_char}
            }
            paragraphs.append(paragraph)
        
        return paragraphs
    
    def _balance_content_distribution(self, content: str, max_sub_points: int = 4) -> List[str]:
        """Balance content distribution - split slides with too many sub-points"""
        lines = content.split('\n')
        balanced_chunks = []
        current_chunk = []
        current_main_point = None
        sub_point_count = 0
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            if line.startswith("● "):
                # New main point
                if current_chunk and sub_point_count > 0:
                    # Save current chunk if it has content
                    balanced_chunks.append('\n'.join(current_chunk))
                    current_chunk = []
                    sub_point_count = 0
                
                current_main_point = line
                current_chunk = [line]
                
            elif "○" in line:
                # Sub point
                if sub_point_count >= max_sub_points:
                    # Too many sub-points, split here
                    balanced_chunks.append('\n'.join(current_chunk))
                    current_chunk = [current_main_point, line] if current_main_point else [line]
                    sub_point_count = 1
                else:
                    current_chunk.append(line)
                    sub_point_count += 1
            else:
                # Regular content
                current_chunk.append(line)
        
        # Add final chunk
        if current_chunk:
            balanced_chunks.append('\n'.join(current_chunk))
        
        return balanced_chunks
    
    def _paste_cover_slide(self, slide: Dict, date: str):
        """Paste date vào cover slide"""
        shapes = slide.get("shapes", [])
        if len(shapes) >= 1:
            shape = shapes[0]
            if "paragraphs" in shape and shape["paragraphs"]:
                paragraph = shape["paragraphs"][0]
                if "runs" in paragraph and len(paragraph["runs"]) >= 2:
                    # Run 1: Date, Run 2: Title
                    paragraph["runs"][0]["text"] = date
                    paragraph["runs"][1]["text"] = "進捗報告書"
    
    def _paste_section_divider(self, slide: Dict, number: str, title: str):
        """Paste section number và title vào section divider"""
        shapes = slide.get("shapes", [])
        for shape in shapes:
            if shape.get("type") == "text":
                pos_x = shape.get("position", {}).get("x")
                
                if shape.get("paragraphs") and shape["paragraphs"][0].get("runs"):
                    if pos_x == 2542873:  # Title shape
                        shape["paragraphs"][0]["runs"][0]["text"] = title
                    elif pos_x == 5151175:  # Number shape  
                        shape["paragraphs"][0]["runs"][0]["text"] = number
    
    def save_script(self, slides: List[Dict[str, Any]]) -> str:
        """Save generated script to JSON file"""
        output_file = self.output_path / "slide_script.json"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(slides, f, ensure_ascii=False, indent=2)
        
        return str(output_file)
    
    def run(self):
        """Main execution method - FUNCTION CALLING approach only"""
        print("📋 Using FUNCTION CALLING approach to generate slides...")
        
        # Generate slides using Function Calling (fast & reliable)
        slides = self.generate_from_structured_format("output/structured_format.txt", "output/slide_script.json")
        
        print(f"Generated {len(slides)} slides")
        
        output_path = self.save_script(slides)
        print(f"Script saved to: {output_path}")
        
        return slides

if __name__ == "__main__":
    generator = SlideScriptGenerator()
    generator.run()