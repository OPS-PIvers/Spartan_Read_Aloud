import re

def extract_function_range(source, func_name):
    start_marker = f"function {func_name}"
    start_index = source.find(start_marker)
    if start_index == -1:
        raise ValueError(f"Function {func_name} not found")

    # Check for JSDoc comment preceding the function
    # Search backwards from start_index for "/**"
    comment_start = -1
    for i in range(start_index - 1, -1, -1):
        if source[i:i+3] == '/**':
            # Check if this comment belongs to the function (only whitespace between them)
            if source[i+3:start_index].strip() == '':
                comment_start = i
            break
        elif source[i] == '}' or source[i] == ';':
            # Hit previous code block, stop searching
            break

    effective_start = comment_start if comment_start != -1 else start_index

    brace_count = 0
    end_index = -1
    found_start = False

    for i in range(start_index, len(source)):
        if source[i] == '{':
            brace_count += 1
            found_start = True
        elif source[i] == '}':
            brace_count -= 1
            if found_start and brace_count == 0:
                end_index = i + 1
                break

    if end_index == -1:
        raise ValueError(f"Could not find end of function {func_name}")

    return effective_start, end_index

with open('Code.js', 'r') as f:
    content = f.read()

try:
    start, end = extract_function_range(content, 'parseHtmlToChunks')
    # Remove the function and surrounding whitespace if needed
    new_content = content[:start].rstrip() + '\n\n' + content[end:].lstrip()

    with open('Code.js', 'w') as f:
        f.write(new_content)
    print("Function parseHtmlToChunks removed successfully.")
except Exception as e:
    print(f"Error: {e}")
