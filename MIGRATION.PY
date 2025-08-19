import re
from pathlib import Path

def convert_mysql_to_postgresql(input_file, output_file):
    # Define all replacements as (pattern, replacement) tuples
    replacements = [
        # Remove MySQL-specific table options
        (r'ENGINE=\w+', ''),
        (r'DEFAULT CHARSET=\w+', ''),
        (r'COLLATE=\w+', ''),
        (r'ROW_FORMAT=\w+', ''),
        (r'STATS_PERSISTENT=\d', ''),
        (r'TABLESPACE=\w+', ''),
        (r'AUTO_INCREMENT=\d+', ''),
        (r'/\*\!.*?\*/', ''),  # Remove conditional comments
        (r'UNSIGNED', ''),
        
        # Data type conversions
        (r'TINYINT\(1\)', 'BOOLEAN'),
        (r'INT\(\d+\)', 'INTEGER'),
        (r'DATETIME', 'TIMESTAMP'),
        (r'DOUBLE\(\d+,\d+\)', 'DOUBLE PRECISION'),
        (r'VARCHAR\(\d+\) CHARACTER SET \w+', 'TEXT'),
        (r'LONGTEXT', 'TEXT'),
        (r'MEDIUMTEXT', 'TEXT'),
        (r"ENUM\('(.*?)'\)", r'TEXT CHECK (VALUE IN (\'\1\'))'),
        
        # Backticks to quotes
        (r'`', '"'),
        
        # Remove LOCK/UNLOCK statements
        (r'LOCK TABLES.*?WRITE;', ''),
        (r'UNLOCK TABLES;', ''),
        
        # Remove USE statements
        (r'^USE .*?;', ''),
        
        # Fix AUTO_INCREMENT
        (r'AUTO_INCREMENT', 'GENERATED ALWAYS AS IDENTITY'),
        
        # Fix INSERT statements
        (r'INSERT INTO "\w+" VALUES', lambda m: m.group(0).replace('VALUES', '(id, ...) VALUES')),

        (r"COMMENT\s*=\s*'[^']*'", ""),
        (r"COMMENT\s+'[^']*'", ""),
        (r",\s*COMMENT\s*=\s*'[^']*'", ""),
        (r"ENGINE\s*=\s*\w+\s+COMMENT\s*=\s*'[^']*'", ""),
        
        # Remove comentários condicionais do MySQL
        (r"/\*\![0-9]*.*?\*/", ""),
        
        # Remove COMMENT como cláusula isolada
        (r"^\s*COMMENT\s+ON.*?;", ""),
        (r"\bCOMMENT\b", ""),
        
        # Remove qualquer linha que contenha apenas COMMENT
        (r"^.*COMMENT.*$", ""),
        
        # Remove espaços extras deixados pelas remoções
        (r"\s+\)", ")"),
        (r",\s*\)", ")"),
        (r"\(\s+", "("),
    ]

    try:
        # Try different encodings
        encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
        content = None
        
        for encoding in encodings:
            try:
                with open(input_file, 'r', encoding=encoding) as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        
        if content is None:
            raise ValueError("Failed to decode file with tested encodings")

        # Apply all replacements
        for pattern, replacement in replacements:
            content = re.sub(pattern, replacement, content, flags=re.IGNORECASE)

        # Post-processing for multi-line statements
        content = re.sub(r',\s*\)', ')', content)  # Fix trailing commas
        content = re.sub(r';\s*;', ';', content)  # Remove duplicate semicolons

        # Write output
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(content)
            
        print(f"Conversion successful! Output saved to {output_file}")
        
    except Exception as e:
        print(f"Error during conversion: {str(e)}")

# Example usage
input_file = 'dump.sql'
output_file = 'database.sql'

convert_mysql_to_postgresql(input_file, output_file)