#!/bin/bash
# Extract markdown cells and section headers from notebook JSON

for notebook in GeoPoll_questionnaire_validation.ipynb newQ_test.ipynb newQ_testv2.ipynb; do
    echo "==================================================="
    echo "FILE: $notebook"
    echo "==================================================="
    
    # Count cells by type
    total=$(grep -c '"cell_type"' "$notebook")
    code=$(grep -c '"cell_type": "code"' "$notebook")
    md=$((total/2 - code))
    
    echo "Total cells: $code code, $md markdown"
    echo ""
    
    # Extract imports and function definitions using different approach
    awk '
        /import|from/ && prev_type=="code" { 
            print "  IMPORT: " substr($0, 1, 80)
        }
        /^def|^class/ && prev_type=="code" { 
            print "  FUNCTION/CLASS: " substr($0, 1, 80)
        }
        /"cell_type": "markdown"/ { 
            prev_type="markdown" 
        }
        /"cell_type": "code"/ { 
            prev_type="code" 
        }
    ' "$notebook" 2>/dev/null || echo "Parsing markdown..."
    
    echo ""
done
