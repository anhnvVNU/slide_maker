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
    
    def _create_few_shot_examples(self) -> str:
        """Create few-shot examples from templates"""
        examples = []
        
        # Example 1: Cover slide  
        examples.append({
            "context": "Create a cover slide for 進捗報告書 dated August 15, 2025",
            "slide": {
                "slide_type": "cover",
                "shapes": [
                    {
                        "type": "text",
                        "position": {"x": 698942, "y": 1734594},
                        "size": {"width": 5741100, "height": 1674300},
                        "paragraphs": [
                            {
                                "runs": [
                                    {
                                        "text": "20250815",
                                        "font": {"name": "Noto Sans JP", "size_pt": 36.0, "bold": True, "color": [0, 0, 0]}
                                    },
                                    {
                                        "text": "進捗報告書",
                                        "font": {"name": "Noto Sans JP", "size_pt": 36.0, "bold": True, "color": [0, 0, 0]}
                                    }
                                ],
                                "alignment": "left",
                                "line_spacing": {"value": 1.5, "type": "points"}
                            }
                        ]
                    }
                ]
            }
        })
        
        # Example 2: Agenda slide
        examples.append({
            "context": "Create agenda slide with fixed Japanese agenda items タスクの目的, 完了済み作業, 次の作業内容",
            "slide": {
                "slide_type": "agenda_slide",
                "shapes": [
                    {
                        "type": "text",
                        "position": {"x": 713225, "y": 2292038},
                        "size": {"width": 3416400, "height": 396600},
                        "paragraphs": [
                            {
                                "runs": [{"text": "アジェンダ", "font": {"name": "Noto Sans JP", "size_pt": 22.5, "bold": True, "color": [65, 65, 67]}}],
                                "alignment": "left"
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "position": {"x": 4291175, "y": 1551650},
                        "size": {"width": 4326900, "height": 1342200},
                        "paragraphs": [
                            {
                                "runs": [{"text": "タスクの目的", "font": {"name": "Noto Sans JP", "size_pt": 17.0, "bold": True, "color": [255, 255, 255]}}],
                                "alignment": "left",
                                "line_spacing": {"value": 1.5, "type": "points"},

                                "bullet": {"type": "numbered", "style": "arabicPeriod"}
                            },
                            {
                                "runs": [{"text": "完了済み作業", "font": {"name": "Noto Sans JP", "size_pt": 17.0, "bold": True, "color": [255, 255, 255]}}],
                                "alignment": "left",
                                "line_spacing": {"value": 1.5, "type": "points"},

                                "bullet": {"type": "numbered", "style": "arabicPeriod"}
                            },
                            {
                                "runs": [{"text": "次の作業内容", "font": {"name": "Noto Sans JP", "size_pt": 17.0, "bold": True, "color": [255, 255, 255]}}],
                                "alignment": "left",
                                "line_spacing": {"value": 1.5, "type": "points"},

                                "bullet": {"type": "numbered", "style": "arabicPeriod"}
                            }
                        ]
                    }
                ]
            }
        })
        
        # Example 3: Section divider
        examples.append({
            "context": "Create section divider for section 01 タスクの目的",
            "slide": {
                "slide_type": "section_divider",
                "shapes": [
                    {
                        "type": "text",
                        "position": {"x": 2542873, "y": 2571750},
                        "size": {"width": 5665500, "height": 862200},
                        "paragraphs": [
                            {
                                "runs": [{"text": "タスクの目的", "font": {"name": "Noto Sans JP", "size_pt": 30.0, "bold": True, "color": [0, 0, 0]}}],
                                "alignment": "right",
                                "line_spacing": {"value": 1.0, "type": "points"}
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "position": {"x": 5151175, "y": 1768609},
                        "size": {"width": 3057300, "height": 862200},
                        "paragraphs": [
                            {
                                "runs": [{"text": "01", "font": {"name": "Noto Sans JP", "size_pt": 60.0, "bold": True, "color": [0, 0, 0]}}],
                                "alignment": "right",
                                "line_spacing": {"value": 1.0, "type": "points"}
                            }
                        ]
                    }
                ]
            }
        })
        
        # Example 4: Content slide
        examples.append({
            "context": "Create タスクの目的 content slide with task overview bullet points",
            "slide": {
                "slide_type": "content",
                "shapes": [
                    {
                        "type": "text",
                        "position": {"x": 584875, "y": 427349},
                        "size": {"width": 7717500, "height": 421500},
                        "paragraphs": [
                            {
                                "runs": [{"text": "1.\tタスクの目的", "font": {"name": "Noto Sans JP", "size_pt": 18.0, "bold": True, "color": [0, 0, 0]}}],
                                "alignment": "left",
                                "level": 0,
                                "line_spacing": {"value": 1.0, "type": "points"}
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "position": {"x": 626250, "y": 923775},
                        "size": {"width": 7891500, "height": 3873600},
                        "paragraphs": [
                            {
                                "runs": [{"text": "AIREAD_ARISE-4149 【AIRead】Trong folder Temp của USERPROFILE còn tồn tại onnxruntime-java", "font": {"name": "Noto Sans JP", "size_pt": 14.0}}],
                                "alignment": "left",
                                "level": 0,
                                "line_spacing": {"value": 1.15, "type": "points"},

                                "bullet": {"type": "bullet", "char": "●"}
                            },
                            {
                                "runs": [{"text": "Bổ sung hàm xóa các file temp tự động", "font": {"name": "Noto Sans JP", "size_pt": 14.0}}],
                                "alignment": "left",
                                "level": 1,
                                "line_spacing": {"value": 1.15, "type": "points"},

                                "bullet": {"type": "bullet", "char": "○"}
                            },
                            {
                                "runs": [{"text": "AIREAD_ARISE-4157 【multi3_jpn】Cải thiện độ chính xác", "font": {"name": "Noto Sans JP", "size_pt": 14.0}}],
                                "alignment": "left",
                                "level": 0,
                                "line_spacing": {"value": 1.15, "type": "points"},

                                "bullet": {"type": "bullet", "char": "●"}
                            },
                            {
                                "runs": [{"text": "Điều tra lỗi và sửa đổi hàm tiền xử lý của multi3_jpn", "font": {"name": "Noto Sans JP", "size_pt": 14.0}}],
                                "alignment": "left",
                                "level": 1,
                                "line_spacing": {"value": 1.15, "type": "points"},

                                "bullet": {"type": "bullet", "char": "○"}
                            },
                            {
                                "runs": [{"text": "AIREAD_ARISE-4034【AIRead】位置合わせ機能の強化: tăng độ chính xác của alignment function", "font": {"name": "Noto Sans JP", "size_pt": 14.0}}],
                                "alignment": "left",
                                "level": 0,
                                "line_spacing": {"value": 1.15, "type": "points"},

                                "bullet": {"type": "bullet", "char": "●"}
                            },
                            {
                                "runs": [{"text": "Sửa ảnh template thành ảnh trước khi tiền xử lý", "font": {"name": "Noto Sans JP", "size_pt": 14.0}}],
                                "alignment": "left",
                                "level": 1,
                                "line_spacing": {"value": 1.15, "type": "points"},

                                "bullet": {"type": "bullet", "char": "○"}
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "position": {"x": 8768388, "y": 4835723},
                        "size": {"width": 283200, "height": 246300},
                        "paragraphs": [
                            {
                                "runs": [{"text": "3", "font": {"name": "Arial", "size_pt": 10.0, "bold": False, "italic": False, "underline": False, "color": [0, 0, 0]}}],
                                "alignment": "left",
                                "level": 0,
                                "line_spacing": {"value": 1.0, "type": "points"}
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "position": {"x": 2653144, "y": -3003},
                        "size": {"width": 3816900, "height": 421500},
                        "fill": {"type": "solid", "color": [255, 236, 185]},
                        "paragraphs": [
                            {
                                "runs": [{"text": "1. タスクの目的", "font": {"name": "Noto Sans JP", "size_pt": 20.0, "bold": True, "italic": False, "underline": False, "color": [0, 0, 0]}}],
                                "alignment": "center",
                                "level": 0,
                                "line_spacing": {"value": 1.0, "type": "points"}
                            }
                        ]
                    }
                ]
            }
        })
        
        # Example 5: Task subtitle - using EXACT subtitle template structure
        examples.append({
            "context": "Create subtitle for task 2.1 AIREAD_ARISE-4149",
            "slide": {
                "slide_type": "subtitle",
                "shapes": [
                    {
                        "type": "text",
                        "position": {"x": 1191600, "y": 1483850},
                        "size": {"width": 6760800, "height": 1555800},
                        "paragraphs": [
                            {
                                "runs": [{"text": "2.1. AIREAD_ARISE-4149 【AIRead】", "font": {"name": "Noto Sans JP", "size_pt": 24.0, "bold": True, "color": [0, 0, 0]}}],
                                "alignment": "center",
                                "line_spacing": {"value": 1.5, "type": "points"}
                            },
                            {
                                "runs": [{"text": "Trong folder Temp của USERPROFILE còn tồn tại onnxruntime-java", "font": {"name": "Noto Sans JP", "size_pt": 24.0, "bold": True, "color": [0, 0, 0]}}],
                                "alignment": "center",
                                "line_spacing": {"value": 1.5, "type": "points"}
                            }
                        ]
                    }
                ]
            }
        })
        
        return json.dumps(examples, ensure_ascii=False, indent=2)
    
    def read_representation(self) -> str:
        """Read the representation from output file"""
        represent_file = self.output_path / "represent.txt"
        with open(represent_file, 'r', encoding='utf-8') as f:
            return f.read()
    
    def generate_presentation_script(self) -> List[Dict[str, Any]]:
        """Generate complete presentation script from representation"""
        
        # Read the representation
        representation = self.read_representation()
        
        # Also read structured_format.txt if exists (contains sub-points)
        structured_format = ""
        structured_file = Path("output/structured_format.txt")
        if structured_file.exists():
            with open(structured_file, 'r', encoding='utf-8') as f:
                structured_format = f.read()
                print("📋 Found structured_format.txt with sub-points - including in AI prompt")
        
        # Get few-shot examples
        few_shot_examples = self._create_few_shot_examples()
        
        # Prepare the prompt
        system_prompt = self.settings['prompts']['slide_generator']['system']
        
        # If structured_format exists, enhance the prompt with sub-points instruction
        if structured_format:
            enhanced_instruction = """
            
IMPORTANT: Use the structured_format.txt below which contains detailed sub-points for slide 4 (タスクの目的). 
Make sure to include ALL sub-points (○ bullet points) in the Purpose slide content.

STRUCTURED FORMAT WITH SUB-POINTS:
{structured_format}

When generating slide 4 (Purpose slide), extract the task list WITH sub-points from the structured format above."""
            
            user_prompt = self.settings['prompts']['slide_generator']['user'].format(
                representation=representation + enhanced_instruction.format(structured_format=structured_format),
                templates=json.dumps(self.slide_templates, ensure_ascii=False, indent=2),
                examples=few_shot_examples
            )
        else:
            user_prompt = self.settings['prompts']['slide_generator']['user'].format(
                representation=representation,
                templates=json.dumps(self.slide_templates, ensure_ascii=False, indent=2),
                examples=few_shot_examples
            )
        
        # Generate slides using OpenAI
        response = self.client.chat.completions.create(
            model=self.settings['openai']['model'],
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=self.settings['openai']['temperature'],
            max_tokens=8192,  # Increased from default to handle large responses
            response_format={"type": "json_object"}
        )
        
        # Parse response with error handling
        response_content = response.choices[0].message.content
        print(f"AI Response length: {len(response_content)} chars")
        
        try:
            result = json.loads(response_content)
        except json.JSONDecodeError as e:
            print(f"JSON Parse Error: {e}")
            print(f"Response content (last 500 chars):")
            print(response_content[-500:])
            
            # Try to fix common JSON issues
            # Remove any trailing text after the JSON array
            try:
                # Find the last closing bracket
                last_bracket = response_content.rfind(']')
                if last_bracket != -1:
                    fixed_content = response_content[:last_bracket+1]
                    result = json.loads(fixed_content)
                    print("Fixed JSON by truncating trailing content")
                else:
                    raise e
            except:
                print("Could not fix JSON. Saving raw response to debug file.")
                with open(self.output_path / "debug_response.txt", 'w', encoding='utf-8') as f:
                    f.write(response_content)
                raise e
        
        # Ensure we have a list of slides
        if isinstance(result, dict) and "slides" in result:
            slides = result["slides"]
        elif isinstance(result, list):
            slides = result
        else:
            slides = [result]
        
        return slides
    
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
        """Main execution method"""
        print("Checking for structured_format.txt...")
        
        structured_file = Path("output/structured_format.txt")
        if structured_file.exists():
            print("📋 Found structured_format.txt - using FUNCTION CALLING approach")
            slides = self.generate_from_structured_format("output/structured_format.txt", "output/slide_script.json")
        else:
            print("📋 No structured_format.txt found - falling back to AI generation")
            print("Reading representation from represent.txt...")
            print("Generating presentation script...")
            slides = self.generate_presentation_script()
        
        print(f"Generated {len(slides)} slides")
        
        output_path = self.save_script(slides)
        print(f"Script saved to: {output_path}")
        
        return slides

if __name__ == "__main__":
    generator = SlideScriptGenerator()
    generator.run()