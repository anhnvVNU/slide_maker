#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Content Formatter: Bước 2 - AI Parse represent.txt và tạo structured format (.txt)
Sử dụng AI để parse represent.txt thành format chuẩn cho script_generator.py
"""

import re
import os
import json
import yaml
from pathlib import Path
from typing import Dict, List, Any
from datetime import datetime
from openai import OpenAI
from dotenv import load_dotenv

class ContentFormatter:
    def __init__(self, settings_path: str = "settings.yaml"):
        # Load environment variables and settings
        load_dotenv(Path(__file__).parent.parent / '.env')
        
        self.settings_path = Path(__file__).parent.parent / settings_path
        with open(self.settings_path, 'r', encoding='utf-8') as f:
            self.settings = yaml.safe_load(f)
        
        self.client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    
    def parse_represent_file(self, represent_path: str) -> Dict[str, Any]:
        """AI Parse represent.txt thành structured data"""
        with open(represent_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print("🤖 Using AI to parse represent.txt...")
        
        # Create AI prompt for parsing
        system_prompt = """You are a technical document parser. Parse the Vietnamese technical report and extract structured information in JSON format.

IMPORTANT: Return ONLY valid JSON, no additional text or explanations.

Required JSON structure:
{
  "date_formatted": "YYYYMMDD format from report date",
  "report_time": "original time period string",
  "executive_summary": "extracted summary text",
  "tasks": [
    {
      "task_number": 1,
      "title": "extracted task title with ID code",
      "problem_description": "what problem was being solved",
      "solution_steps": [
        {"text": "step description", "level": 0 or 1},
        {"text": "sub-step description", "level": 1}
      ],
      "result": "what was achieved",
      "analysis": "analysis of the work", 
      "comment": "additional comments"
    }
  ],
  "future_work": "next work to be done"
}

Instructions:
- Extract date from "Thời gian:" and format as YYYYMMDD
- Get summary from "TÓM TẮT ĐIỀU HÀNH:" section
- Parse each "NHIỆM VỤ X:" as separate task
- For solution_steps: level 0 for main points (▪), level 1 for sub-points
- Get future work from "CÔNG VIỆC SẮP TỚI:" section"""

        user_prompt = f"""Parse this Vietnamese technical report:

{content}

Return the parsed data in the exact JSON structure specified."""

        # Call OpenAI API
        response = self.client.chat.completions.create(
            model=self.settings['openai']['model'],
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.1,  # Low temperature for consistent parsing
            max_tokens=4000,
            response_format={"type": "json_object"}
        )
        
        # Parse AI response
        response_content = response.choices[0].message.content
        print(f"AI Response length: {len(response_content)} chars")
        
        try:
            parsed_data = json.loads(response_content)
            print(f"✓ AI parsed {len(parsed_data.get('tasks', []))} tasks")
            return parsed_data
        except json.JSONDecodeError as e:
            print(f"❌ JSON Parse Error: {e}")
            print(f"Response content: {response_content[:500]}...")
            # Fallback to empty structure
            return {
                "date_formatted": datetime.now().strftime("%Y%m%d"),
                "report_time": "N/A",
                "executive_summary": "",
                "tasks": [],
                "future_work": ""
            }
    

    
    def create_structured_format(self, parsed_data: Dict[str, Any]) -> str:
        """Tạo structured format (.txt) với THỨ TỰ CỐ ĐỊNH theo settings"""
        
        # Calculate slide numbers theo ĐÚNG THỨ TỰ - count unique tasks only
        unique_task_titles = set()
        for task in parsed_data['tasks']:
            unique_task_titles.add(task['title'])
        total_tasks = len(unique_task_titles)
        
        structured_content = f"""=== SLIDE STRUCTURE FORMAT (THỨ TỰ CỐ ĐỊNH) ===
Report Date: {parsed_data['date_formatted']}
Report Period: {parsed_data['report_time']}
Total Tasks: {total_tasks}

=== SLIDE ORGANIZATION (FIXED ORDER) ===

SLIDE 1: Cover
- Type: cover
- Content: "{parsed_data['date_formatted']}進捗報告書"

SLIDE 2: Agenda  
- Type: agenda_slide
- Content: Fixed agenda items (CỐ ĐỊNH)
  1. タスクの目的
  2. 完了済み作業  
  3. 次の作業内容

SLIDE 3: Section Divider 01
- Type: section_divider
- Number: "01"
- Title: "タスクの目的"

SLIDE 4: Purpose Overview
- Type: content
- Title: "1. タスクの目的"
- Content: Task list overview
"""

        # Add task list với bullet points (loại bỏ trùng lặp) + sub-points
        seen_tasks = set()
        unique_tasks = []
        for task in parsed_data['tasks']:
            if task['title'] not in seen_tasks:
                unique_tasks.append(task)
                seen_tasks.add(task['title'])
        
        # Build task list with sub-points
        for task in unique_tasks:
            structured_content += f"  ● {task['title']}\n"
            # Add 1-2 brief sub-points for each main task
            if task.get('problem_description'):
                brief_desc = task['problem_description'][:100] + "..." if len(task['problem_description']) > 100 else task['problem_description']
                structured_content += f"    ○ {brief_desc}\n"
            if task.get('result'):
                brief_result = task['result'][:100] + "..." if len(task['result']) > 100 else task['result']
                structured_content += f"    ○ Kết quả: {brief_result}\n"
        
        structured_content += f"""
SLIDE 5: Section Divider 02
- Type: section_divider  
- Number: "02"
- Title: "完了済み作業"

"""

        # Task slides: SUBTITLE + 3 CONTENT slides cho mỗi task
        current_slide = 6
        for task_idx, task in enumerate(parsed_data['tasks'], 1):
            
            # Subtitle slide cho task
            structured_content += f"""SLIDE {current_slide}: Task {task_idx} Subtitle
- Type: subtitle
- Content: "2.{task_idx}. {task['title']}"

"""
            current_slide += 1
            
            # Content slide 1: Problem description
            structured_content += f"""SLIDE {current_slide}: Task {task_idx} Problem Description
- Type: content
- Title: "2. 完了済み作業"
- Content:
  ● {task['problem_description']}

"""
            current_slide += 1
            
            # Content slide 2: Solution steps
            structured_content += f"""SLIDE {current_slide}: Task {task_idx} Solution Process
- Type: content
- Title: "2. 完了済み作業"
- Content: Solution steps
"""
            for step in task['solution_steps']:
                bullet = "●" if step['level'] == 0 else "○"
                structured_content += f"  {bullet} {step['text']}\n"
            structured_content += "\n"
            current_slide += 1
            
            # Content slide 3: Results & Analysis
            structured_content += f"""SLIDE {current_slide}: Task {task_idx} Results & Analysis
- Type: content
- Title: "2. 完了済み作業"
- Content:
  ● {task['result']}
  ● {task['analysis']}
  ● {task['comment']}

"""
            current_slide += 1
        
        # Final slides: Section divider + Future work
        section_03_slide = current_slide
        future_work_slide = current_slide + 1
        
        structured_content += f"""SLIDE {section_03_slide}: Section Divider 03
- Type: section_divider
- Number: "03"  
- Title: "次の作業内容"

SLIDE {future_work_slide}: Future Work
- Type: content
- Title: "3. 次の作業内容"
- Content:
  ● {parsed_data['future_work']}

=== SLIDE ORDER VERIFICATION ===
1. Cover
2. Agenda  
3. Section 01 ("01" + "タスクの目的")
4. Content ("タスクの目的" với task list only)
5. Section 02 ("02" + "完了済み作業")
6-{current_slide-1}. Tasks (mỗi task = subtitle + 3 content slides)
{section_03_slide}. Section 03 ("03" + "次の作業内容")
{future_work_slide}. Future Work content

=== END STRUCTURE ===
Total Slides: {future_work_slide}
Expected Pattern: Cover → Agenda → Sec01 → Purpose → Sec02 → [Task blocks] → Sec03 → Future
"""
        
        return structured_content
    
    def run(self, represent_path: str = "output/represent.txt", 
            output_path: str = "output/structured_format.txt") -> str:
        """Main function: Parse và tạo structured format"""
        
        print("🔍 Step 2: Parsing represent.txt...")
        parsed_data = self.parse_represent_file(represent_path)
        print(f"✓ Found {len(parsed_data['tasks'])} tasks")
        
        print("📝 Creating structured format...")
        structured_content = self.create_structured_format(parsed_data)
        
        # Save to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(structured_content)
        
        print(f"✓ Structured format saved to {output_path}")
        return structured_content

if __name__ == "__main__":
    formatter = ContentFormatter()
    structured_content = formatter.run()
