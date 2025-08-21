import sys
import re
from collections import Counter

def count_tokens_simple(text):
    """
    Simple token counter using whitespace and punctuation splitting
    Similar to how most LLMs tokenize text
    """
    # Split by whitespace and common punctuation
    tokens = re.findall(r'\b\w+\b|[^\w\s]', text)
    return len(tokens)

def count_tokens_detailed(text):
    """
    More detailed token analysis
    """
    # Different token types
    words = re.findall(r'\b\w+\b', text)
    numbers = re.findall(r'\b\d+\b', text)
    punctuation = re.findall(r'[^\w\s]', text)
    
    # Character and line counts
    char_count = len(text)
    line_count = text.count('\n') + 1
    
    # Word frequency
    word_freq = Counter(words)
    
    return {
        'total_tokens': len(words) + len(punctuation),
        'word_tokens': len(words),
        'unique_words': len(set(words)),
        'number_tokens': len(numbers),
        'punctuation_tokens': len(punctuation),
        'characters': char_count,
        'lines': line_count,
        'average_word_length': sum(len(w) for w in words) / len(words) if words else 0,
        'most_common_words': word_freq.most_common(10)
    }

def estimate_openai_tokens(text):
    """
    Rough estimate of OpenAI GPT tokens
    OpenAI's tokenizer typically creates ~1.3 tokens per word in English
    For mixed languages (Japanese/English), it's typically higher
    """
    words = len(re.findall(r'\b\w+\b', text))
    
    # Check for Japanese characters
    has_japanese = bool(re.search(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]', text))
    
    if has_japanese:
        # Japanese text typically uses more tokens
        estimated_tokens = int(words * 2.5)
    else:
        estimated_tokens = int(words * 1.3)
    
    return estimated_tokens

def analyze_file(filename):
    """
    Analyze token counts in a file
    """
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            text = f.read()
        
        print(f"\n{'='*60}")
        print(f"TOKEN ANALYSIS FOR: {filename}")
        print('='*60)
        
        # Simple token count
        simple_count = count_tokens_simple(text)
        print(f"\nSimple Token Count: {simple_count:,}")
        
        # Detailed analysis
        detailed = count_tokens_detailed(text)
        print(f"\nDetailed Analysis:")
        print(f"  Total Tokens: {detailed['total_tokens']:,}")
        print(f"  Word Tokens: {detailed['word_tokens']:,}")
        print(f"  Unique Words: {detailed['unique_words']:,}")
        print(f"  Number Tokens: {detailed['number_tokens']:,}")
        print(f"  Punctuation Tokens: {detailed['punctuation_tokens']:,}")
        print(f"  Total Characters: {detailed['characters']:,}")
        print(f"  Total Lines: {detailed['lines']:,}")
        print(f"  Average Word Length: {detailed['average_word_length']:.2f} characters")
        
        # Estimated OpenAI tokens
        openai_estimate = estimate_openai_tokens(text)
        print(f"\nEstimated OpenAI Tokens: ~{openai_estimate:,}")
        print(f"  (This is a rough estimate. Actual count may vary)")
        
        # Most common words
        print(f"\nTop 10 Most Common Words:")
        for word, count in detailed['most_common_words']:
            print(f"  '{word}': {count} times")
        
        # File size
        import os
        file_size = os.path.getsize(filename)
        print(f"\nFile Size: {file_size:,} bytes ({file_size/1024:.2f} KB)")
        
        # Token density
        token_density = detailed['total_tokens'] / detailed['lines']
        print(f"Token Density: {token_density:.2f} tokens per line")
        
        return detailed
        
    except FileNotFoundError:
        print(f"Error: File '{filename}' not found")
        return None
    except Exception as e:
        print(f"Error processing file: {e}")
        return None

def main():
    if len(sys.argv) > 1:
        filename = sys.argv[1]
    else:
        filename = 'slide_output_enhanced.txt'
    
    analyze_file(filename)

if __name__ == "__main__":
    main()