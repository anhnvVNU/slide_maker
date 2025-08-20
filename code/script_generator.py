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
        
        # Get few-shot examples
        few_shot_examples = self._create_few_shot_examples()
        
        # Prepare the prompt
        system_prompt = self.settings['prompts']['slide_generator']['system']
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
            max_tokens=self.settings['openai']['max_tokens'],
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
    
    def save_script(self, slides: List[Dict[str, Any]]) -> str:
        """Save generated script to JSON file"""
        output_file = self.output_path / "slide_script.json"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(slides, f, ensure_ascii=False, indent=2)
        
        return str(output_file)
    
    def run(self):
        """Main execution method"""
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