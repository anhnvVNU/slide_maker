#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Main orchestrator for the slide maker pipeline
Runs content extraction, script generation, and template building in sequence
"""

import os
import sys
import json
import time
from pathlib import Path
from typing import Dict, Any
import logging

# Add the code directory to the path to import modules
current_dir = Path(__file__).parent
code_dir = current_dir / "code"
sys.path.insert(0, str(code_dir))

# Import the modules
from content_extractor import ContentExtractor
from content_formatter import ContentFormatter
from script_generator import SlideScriptGenerator
from master_template_builder import main as build_presentation

def setup_logging() -> logging.Logger:
    """Setup logging configuration"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('output/slide_maker.log'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def check_prerequisites() -> bool:
    """Check if all required files and directories exist"""
    logger = logging.getLogger(__name__)
    
    required_files = [
        "settings.yaml",
        "data/raw_data.txt",
        "template/template.pptx"
    ]
    
    required_dirs = [
        "template/cover_slide",
        "template/agenda_slide", 
        "template/content",
        "template/section_divider",
        "template/subtitle"
    ]
    
    all_exists = True
    
    # Check files
    for file_path in required_files:
        if not Path(file_path).exists():
            logger.error(f"Required file missing: {file_path}")
            all_exists = False
        else:
            logger.info(f"✓ Found: {file_path}")
    
    # Check directories
    for dir_path in required_dirs:
        if not Path(dir_path).exists():
            logger.error(f"Required directory missing: {dir_path}")
            all_exists = False
        else:
            logger.info(f"✓ Found: {dir_path}")
    
    # Create output directory if it doesn't exist
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)
    logger.info(f"✓ Output directory ready: {output_dir}")
    
    return all_exists

def run_content_extraction() -> Dict[str, Any]:
    """Run content extraction step"""
    logger = logging.getLogger(__name__)
    logger.info("=== STEP 1: CONTENT EXTRACTION ===")
    
    try:
        extractor = ContentExtractor()
        
        # Process files from data folder
        data_folder = Path("data")
        results = []
        
        if data_folder.exists():
            for file in data_folder.iterdir():
                if file.is_file() and file.suffix in ['.txt', '.csv', '.xlsx']:
                    logger.info(f"Processing file: {file.name}")
                    result = extractor.process_file(str(file), save_represent=True)
                    results.append(result)
                    logger.info(f"✓ Processed: {file.name}")
        else:
            logger.error("Data folder not found!")
            return {"success": False, "error": "Data folder not found"}
        
        if not results:
            logger.error("No data files found to process!")
            return {"success": False, "error": "No data files found"}
        
        logger.info(f"✓ Content extraction completed. Processed {len(results)} files.")
        logger.info("✓ represent.txt saved to output/")
        
        return {
            "success": True,
            "files_processed": len(results),
            "results": results
        }
        
    except Exception as e:
        logger.error(f"Content extraction failed: {str(e)}")
        return {"success": False, "error": str(e)}

def run_content_formatting() -> Dict[str, Any]:
    """Run content formatting step - AI parse represent.txt to structured_format.txt"""
    logger = logging.getLogger(__name__)
    logger.info("=== STEP 2: CONTENT FORMATTING ===")
    
    try:
        # Check if represent.txt exists
        represent_file = Path("output/represent.txt")
        if not represent_file.exists():
            logger.error("represent.txt not found. Run content extraction first.")
            return {"success": False, "error": "represent.txt not found"}
        
        formatter = ContentFormatter()
        structured_content = formatter.run()
        
        logger.info("✓ Content formatting completed.")
        logger.info("✓ structured_format.txt saved to output/")
        
        return {
            "success": True,
            "structured_content": structured_content
        }
        
    except Exception as e:
        logger.error(f"Content formatting failed: {str(e)}")
        return {"success": False, "error": str(e)}

def run_script_generation() -> Dict[str, Any]:
    """Run script generation step"""
    logger = logging.getLogger(__name__)
    logger.info("=== STEP 3: SCRIPT GENERATION ===")
    
    try:
        # Check if structured_format.txt exists
        structured_file = Path("output/structured_format.txt")
        if not structured_file.exists():
            logger.error("structured_format.txt not found. Run content formatting first.")
            return {"success": False, "error": "structured_format.txt not found"}
        
        generator = SlideScriptGenerator()
        slides = generator.run()
        
        logger.info(f"✓ Script generation completed. Generated {len(slides)} slides.")
        logger.info("✓ slide_script.json saved to output/")
        
        return {
            "success": True,
            "slides_generated": len(slides),
            "slides": slides
        }
        
    except Exception as e:
        logger.error(f"Script generation failed: {str(e)}")
        return {"success": False, "error": str(e)}

def run_presentation_build() -> Dict[str, Any]:
    """Run presentation building step"""
    logger = logging.getLogger(__name__)
    logger.info("=== STEP 4: PRESENTATION BUILDING ===")
    
    try:
        # Check if slide_script.json exists
        script_file = Path("output/slide_script.json")
        if not script_file.exists():
            logger.error("slide_script.json not found. Run script generation first.")
            return {"success": False, "error": "slide_script.json not found"}
        
        # Run the presentation builder
        build_presentation()
        
        # Check if the output file was created
        output_file = Path("output/final_presentation.pptx")
        if output_file.exists():
            logger.info("✓ Presentation building completed.")
            logger.info("✓ final_presentation.pptx saved to output/")
            return {
                "success": True,
                "output_file": str(output_file),
                "file_size": output_file.stat().st_size
            }
        else:
            logger.error("Presentation file was not created.")
            return {"success": False, "error": "Output file not created"}
        
    except Exception as e:
        logger.error(f"Presentation building failed: {str(e)}")
        return {"success": False, "error": str(e)}

def main():
    """Main function to run the complete pipeline"""
    start_time = time.time()
    logger = setup_logging()
    
    logger.info("="*60)
    logger.info("SLIDE MAKER PIPELINE STARTED")
    logger.info("="*60)
    
    # Check prerequisites
    logger.info("Checking prerequisites...")
    if not check_prerequisites():
        logger.error("Prerequisites check failed. Please ensure all required files exist.")
        sys.exit(1)
    
    # Store results for summary
    pipeline_results = {}
    
    # Step 1: Content Extraction
    extraction_result = run_content_extraction()
    pipeline_results["extraction"] = extraction_result
    
    if not extraction_result["success"]:
        logger.error("Pipeline failed at content extraction step.")
        sys.exit(1)
    
    # Step 2: Content Formatting (AI Parse)
    formatting_result = run_content_formatting()
    pipeline_results["formatting"] = formatting_result
    
    if not formatting_result["success"]:
        logger.error("Pipeline failed at content formatting step.")
        sys.exit(1)
    
    # Step 3: Script Generation
    script_result = run_script_generation()
    pipeline_results["script"] = script_result
    
    if not script_result["success"]:
        logger.error("Pipeline failed at script generation step.")
        sys.exit(1)
    
    # Step 4: Presentation Building
    build_result = run_presentation_build()
    pipeline_results["build"] = build_result
    
    if not build_result["success"]:
        logger.error("Pipeline failed at presentation building step.")
        sys.exit(1)
    
    # Success summary
    end_time = time.time()
    total_time = end_time - start_time
    
    logger.info("="*60)
    logger.info("PIPELINE COMPLETED SUCCESSFULLY")
    logger.info("="*60)
    logger.info(f"Total execution time: {total_time:.2f} seconds")
    logger.info("")
    logger.info("SUMMARY:")
    logger.info(f"• Files processed: {extraction_result.get('files_processed', 0)}")
    logger.info(f"• Slides generated: {script_result.get('slides_generated', 0)}")
    logger.info(f"• Output file: {build_result.get('output_file', 'N/A')}")
    
    if build_result.get('file_size'):
        file_size_mb = build_result['file_size'] / (1024 * 1024)
        logger.info(f"• File size: {file_size_mb:.2f} MB")
    
    logger.info("")
    logger.info("Output files:")
    logger.info("• output/represent.txt - Formatted content")
    logger.info("• output/structured_format.txt - AI-parsed structured format")
    logger.info("• output/slide_script.json - Generated slides script")
    logger.info("• output/final_presentation.pptx - Final PowerPoint presentation")
    
    # Save pipeline results
    with open("output/pipeline_results.json", "w", encoding="utf-8") as f:
        pipeline_results["execution_time"] = total_time
        pipeline_results["timestamp"] = time.strftime("%Y-%m-%d %H:%M:%S")
        json.dump(pipeline_results, f, ensure_ascii=False, indent=2)
    
    logger.info("• output/pipeline_results.json - Pipeline execution summary")
    logger.info("="*60)

if __name__ == "__main__":
    main()
