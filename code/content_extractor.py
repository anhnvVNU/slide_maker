import os
import json
from pathlib import Path
import yaml
from openai import OpenAI
from dotenv import load_dotenv
import pandas as pd

load_dotenv()


class ContentExtractor:
    def __init__(self, settings_path: str = "settings.yaml"):
        self.settings = self._load_settings(settings_path)
        self.client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        
    def _load_settings(self, settings_path: str) -> dict:
        with open(settings_path, 'r') as f:
            return yaml.safe_load(f)
    
    def read_file(self, file_path: str) -> str:
        """Read content from txt, csv, or excel files"""
        file_path = Path(file_path)
        
        if file_path.suffix == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        elif file_path.suffix == '.csv':
            df = pd.read_csv(file_path)
            return df.to_string()
        elif file_path.suffix in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
            return df.to_string()
        else:
            raise ValueError(f"Unsupported file type: {file_path.suffix}")
    
    def extract_and_reformat(self, raw_content: str) -> str:
        """
        Extract and reformat raw content into structured text format for slides
        """
        system_prompt = self.settings['prompts']['content_reformatter']['system']
        user_prompt = self.settings['prompts']['content_reformatter']['weekly_report'].format(
            content=raw_content
        )
        
        try:
            response = self.client.chat.completions.create(
                model=self.settings['openai']['model'],
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=self.settings['openai']['temperature'],
                max_tokens=self.settings['openai']['max_tokens']
            )
            
            result = response.choices[0].message.content
            return result
            
        except Exception as e:
            print(f"Error during extraction: {e}")
            return f"Error processing content: {str(e)}"
    
    def process_file(self, file_path: str, save_represent: bool = True) -> dict:
        """Process a file from the data folder"""
        print(f"Reading file: {file_path}")
        raw_content = self.read_file(file_path)
        
        print(f"Extracting and reformatting content...")
        extracted_text = self.extract_and_reformat(raw_content)
        
        # Save represent.txt if requested
        if save_represent:
            output_dir = "output"
            os.makedirs(output_dir, exist_ok=True)
            
            represent_file = f"{output_dir}/represent.txt"
            with open(represent_file, 'w', encoding='utf-8') as f:
                f.write(extracted_text)
            print(f"Saved represent.txt to: {represent_file}")
        
        return {
            'source_file': Path(file_path).name,
            'formatted_content': extracted_text
        }


if __name__ == "__main__":
    # Example usage
    extractor = ContentExtractor()
    
    # Process files from data folder
    data_folder = Path("data")
    if data_folder.exists():
        for file in data_folder.iterdir():
            if file.is_file() and file.suffix in ['.txt', '.csv', '.xlsx']:
                print(f"\nProcessing {file.name}...")
                result = extractor.process_file(str(file))
                print(json.dumps(result, indent=2))