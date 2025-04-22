import os
import sys
import re
import csv
import time
from collections import Counter
from docx2python import docx2python

def clear():
    # command line clear
    os.system('cls' if os.name == 'nt' else 'clear')

def locate_files():
    # locates files for parsing given user inputted folder/file path

    # choose extension type if a folder is given as path
    def extension_type():
        while True:
            extensions = {'1': 'docx', '2': 'txt'}
            ext_choice = input('\n    File extension to target within folder?\n\n    1 = docx, 2 = txt:  ')
            if ext_choice in extensions:
                ext = extensions[ext_choice]
                break
            else:
                clear()
                print('\n    Choice not valid.')
                continue

        return ext

    # scan directory & subdirectories
    def scan_directory(path, ext, parent_folder=''):
        scanned_files = []
        for entry in os.listdir(path):
            entry_path = os.path.join(path, entry)

            if os.path.isdir(entry_path):
                scanned_files.extend(scan_directory(entry_path, ext, os.path.join(parent_folder, entry)))
            elif entry.lower().endswith(ext):
                scanned_files.append([entry, entry_path, parent_folder])
        
        return scanned_files
    
    # returns list with file names and paths
    clear()
    while True:
        path = input('\n    Enter path to file or folder:  ').strip('"\'')
        if os.path.exists(path):
            if os.path.isdir(path):
                clear()
                ext = extension_type()
                files = scan_directory(path, ext)
                if len(files) == 0:
                    clear()
                    print(f'\n    No {ext} files found in this folder.')
                    continue
                break
            else:
                ext = os.path.splitext(path)[1].strip('.')
                if ext not in ('docx', 'txt', 'DOCX', 'TXT'):
                    clear()
                    print('\n    File type must be docx or txt.')
                    continue
                files = [[os.path.basename(path), path, '']]
                break
        else:
            clear()
            print('\n    Path not valid.')
            continue

    return files, ext, path

def options():
    # the below options dictate the regex to be used in the search/what to allow for in IDs, as well as:
        # whether to collapse lines

    clear()
    choices = {'1': True, '2': False}
    while True:
        combine_choice = input(f'\n    Collapse lines when reading text?\n\n    1 = Yes,  2 = No:  ')
        if combine_choice in choices:
            combine = choices[combine_choice]
            break
        else:
            clear()
            print('\n    Choice not valid.')
            continue

    clear()
    choices = {'1': ' ', '2': ''}
    while True:
        whitespace_choice = input(f'\n    Allow for whitespace characters in Doc IDs?\n\n    1 = Yes,  2 = No:  ')
        if whitespace_choice in choices:
            whitespace = choices[whitespace_choice]
            break
        else:
            clear()
            print('\n    Choice not valid.')
            continue        

    if whitespace == ' ':
        clear()
        while True:
            try:
                flex = int(input("\n    Whitespace allowed:  (1 to 5):  "))
                if flex <= 5 and flex >= 1:
                    break
                else:
                    clear()
                    print('\n    Must be 1 to 5, inclusive.')
                    continue
            except:
                clear()
                print('\n    Input not valid.')
                continue
    else:
        flex = 0

    try:
        with open('Suffix.txt', 'r') as suffix_input:
            letters = [l[0] for l in suffix_input.readlines() if l != '']
    except:
        letters = []
        
    if letters != []:
        choices = {'1': 'negative-lookahead-swap', '2': ''}
        clear()
        while True:
            suffix_choice = input(f'\n    Match suffixing letters?\n\n    1 = Yes,  2 = No:  ')
            if suffix_choice in choices:
                suffix = choices[suffix_choice]
                break
            else:
                clear()
                print('\n    Choice not valid.')
                continue
    else:
        suffix = ''

    clear()
    digits = '[1234567890SOoDlIi!|bQ]'
    choices = {'1': '(?:' + digits + '(?:' + whitespace + '{0,' + str(flex) + '}' + digits + '){2,19})+)+' if whitespace == ' ' else digits + '{3,20})+','2': '(?:\d(?:' + whitespace + '{0,' + str(flex) + '}\d){2,19})+)+' if whitespace == ' ' else '\d{3,20})+'}
    boolean_choices = {'1': True, '2': False}

    while True:
        widen_choice = input(f'\n    Allow for character conversion errors?\n\n    1 = Yes,  2 = No:  ')
        if widen_choice in choices:
            widen = choices[widen_choice]
            widen_boolean = boolean_choices[widen_choice]
            break
        else:
            clear()
            print('\n    Choice not valid.')
            continue
      
    return combine, suffix, widen, whitespace, flex, letters, widen_boolean

def read_text(files):
    # reading and storing text from files

    index = -1
    count = 0
    non_combine = ''

    # reading .txt files
    if ext == 'txt':
        for file in files:
            count += 1
            index += 1
            if len(files) > 1:
                clear()
                print(f'\n    Reading file {count}/{len(files)} ...')
            elif os.path.getsize(file[1]) / (1024 ** 2) > 0.15:
                clear()
                print(f"\n    Reading file '{file[0]}' ...")
            try:
                with open(file[1], 'r') as f:
                    non_combine += '\n'.join(f.readlines())
                    non_combine += ' '
                    f.seek(0)
                    if combine:
                        text = ''.join(f.readlines())
                        files[index][1] = text.replace('\n', '')
                    else:
                        files[index][1] = '\n'.join(f.readlines())
            except:
                try:
                    with open(file[1], 'r', encoding='utf-8') as f:
                        non_combine += '\n'.join(f.readlines())
                        non_combine += ' '    
                        f.seek(0)                    
                        if combine:
                            text = ''.join(f.readlines())
                            files[index][1] = text.replace('\n', '')
                        else:
                            files[index][1] = '\n'.join(f.readlines())
                except:
                    clear()
                    input(f"\n    Error decoding '{os.path.basename(file[1])}'.  Hit enter to exit ...")
                    sys.exit()

    else:
        # reading .docx files 
        for file in files:
            count += 1
            index += 1
            if len(files) > 1:
                clear()
                print(f'\n    Reading file {count}/{len(files)} ...')
            elif os.path.getsize(file[1]) / (1024 ** 2) > 0.15:
                clear()
                print(f"\n    Reading file '{file[0]}' ...")
            try:
                with docx2python(file[1]) as content:
                    non_combine += content.text
                    non_combine += ' '
                    if combine:
                        files[index][1] = content.text.replace('\n', '')
                    else:
                        files[index][1] = content.text
            except:
                clear()
                input(f'\n    Error decoding {os.path.basename(file[1])}.  Hit enter to exit ...')
                sys.exit()

        # attempting to structure each word doc footnote-wise; if 1 doc fails, all docs revert to original structure
        structure_error = False
        files_backup = files.copy()
        while True:
            for file in files:  
                try:
                    finds = re.findall(r'footnote\d+\)(?:(?!footnote).)*', file[1])
                    footnotes = []
                    index = -1
                    for find in finds:
                        index += 1
                        footnote_n = int(re.findall('\d+', re.findall('footnote\d+\)', find)[0])[0])
                        entry = []
                        entry.append(f'----footnote{footnote_n}----')
                        entry.append(find.split(f'footnote{footnote_n})')[1].strip())
                        footnotes.append(entry)

                    for footnote in footnotes:
                        file[1] = file[1].replace(footnote[0], ' ' + footnote[1])

                    for find in finds:
                        file[1] = file[1].replace(find, '')
                
                except:
                    clear()
                    structure_error = True
                    input('\n    Error maintaining body/footnote structure.\n\n    All footnote references will appear below all main body references in results.  Hit enter to continue ...')
                    break

                if structure_error:
                    files = files_backup
                    break
            break

        # removes hyperlink references from text
        for file in files:
            try:
                hyperlinks = re.findall(r'<a href=".{5,255}?">', file[1])
                for hyperlink in hyperlinks:
                    file[1] = file[1].replace(hyperlink, '')
            except:
                pass

    # returns files list, except with file paths replaced with texts of documents

    return files, non_combine

def suggest():
    # generating suggested prefixes using advanced regex

    # determine minimum length of any suggested prefixes
    clear()
    print(f'\n    {len(files)} {ext} file(s) loaded.')
    while True:
            min_length = input(f"\n    Enter minimum length for Suggested Prefixes (Enter 's' to skip suggestions):  ")
            if min_length == 's':
                return ''
            else:
                try:
                    min_length = int(min_length)
                    break
                except:
                    clear()
                    print('\n    Input not valid.')
                    continue
    
    # generating analysis text for suggested prefixes (in line with options previously chosen)
    alter = False
    if combine:
        analysis_text = non_combine.replace('\n', '')
        alter = True
    else:
        analysis_text = non_combine

    if whitespace == ' ':
        alter = True
        for i in range(1, flex + 1):
            analysis_text = analysis_text.replace(' ' * i, '')

    # lowercase in suggested prefixes?
    clear()
    choices = {'1': 'a-z', '2': ''}
    while True:
        lowercase_choice = input(f'\n    Allow for lowercase characters in suggested prefixes?\n\n    1 = Yes,  2 = No:  ')
        if lowercase_choice in choices:
            lowercase = choices[lowercase_choice]
            break
        else:
            clear()
            print('\n    Choice not valid.')
            continue

    # initial non-dedupe suggested prefixes
    analysis_text_finds = re.findall('(?:[' + lowercase + 'A-Z]{2,10}+[-_\.]{0,2})+(?:\d[-_\.]{0,2})?[1234567890]{3,15}+', analysis_text)

    index = -1
    for find in analysis_text_finds:
        index += 1
        prefix = ''
        for char in find:
            if not char.isdigit():
                prefix += char
        analysis_text_finds[index] = [prefix.rstrip(''.join(c for c in prefix if not c.isalpha())), find]
    suggested_prefixes = set([f[0] for f in analysis_text_finds])

    # only attempts to cull prefix list if analysis text is altered (whitespaces removed or lines collapsed) and suffixing letters are present, as there is a risk of overlapping prefixes that should be consolidated for searching
    if alter and len(letters) > 0:

        # generates checking list by comparing prefix with all others
        to_check = set()
        for item in suggested_prefixes:
            temp_list = [s for s in suggested_prefixes if s != item]
            for t in temp_list:
                if t.replace(item, '') in letters:
                    to_check.add((t, t.replace(item, '')))

        checked = 0
        times = []
        starting_range = 0
        ending_range = len(to_check)

        to_remove = set()
        to_add = set()
        for check in to_check:
            if ending_range > 6:
                clear()
                print(f'\n    Analysing document(s) for prefixes ...')
                if checked > 0:
                    print(f'\n    Estimated time remaining:  {f"{int(((ending_range - starting_range - checked) * sum(times) / len(times)) // 3600)} hours and {int((((ending_range - starting_range - checked) * sum(times) / len(times)) % 3600) // 60)} minute(s)" if ((ending_range - starting_range - checked) * sum(times) / len(times)) // 3600 >= 1 else f"{int(((ending_range - starting_range - checked) * sum(times) / len(times)) // 60)} minutes and {int(((ending_range - starting_range - checked) * sum(times) / len(times)) % 60)} seconds"}')
            start_time = time.time()
            checked += 1
            
            further_texts = re.findall('.{0,20}?' + check[0], analysis_text)
            for text in further_texts:
                matches = re.findall(r'(?:' + '|'.join(set(suggested_prefixes)) + ')[-_\.]{0,2}[1234567890]{3,15}+' + check[1], text)
                if len(matches) > 0:
                    to_remove.add(check[0])
                else:
                    to_add.add(check[0])

            end_time = time.time()
            times.append(end_time - start_time)

        for r in to_remove:
            suggested_prefixes.remove(r)
        for a in to_add:
            suggested_prefixes.add(a)

    shortened_list = []
    for prefix in suggested_prefixes:
            if len(prefix) >= min_length and len(prefix) <= 20:
                    shortened_list.append(prefix)

    suggested_prefixes = sorted(list(shortened_list), key=lambda x: (len(x), x))

    return suggested_prefixes

def get_prefixes(suggested_prefixes):
    # getting list of prefixes to search, from suggested prefixes and any user inputted prefixes

    clear()
    if suggested_prefixes == []:
        print('\n    Could not generate prefix suggestions.')
    elif suggested_prefixes != '':
            orig_suggested_prefixes = suggested_prefixes.copy()
    
    # user inputted list of prefixes to search for
    prefixes = []
    while True:
        if prefixes != []:
            print('\n    *Accepted  ->  ' + '  |  '.join(prefixes))
        if suggested_prefixes != [] and suggested_prefixes != '':
            print('\n    Suggested Prefixes ->  ' + '  |  '.join(suggested_prefixes))

        p = input('\n    Enter a prefix to search for (Leave empty to continue):  ')
        try:
            if p == '':
                if prefixes == []:
                    clear()
                    print('\n    List cannot be empty.')
                    continue
                else:
                    break
            elif p == "sa":
                for prefix in suggested_prefixes:
                    prefixes.append(prefix)
                suggested_prefixes = []
                clear()
                continue
            elif p == "r":
                if len(suggested_prefixes) > 0:
                    suggested_prefixes.pop(0)
                orig_suggested_prefixes.remove(suggested_prefixes[0])
                clear()
                continue
            elif p == "s":
                if len(suggested_prefixes) > 0:
                    prefixes.append(suggested_prefixes[0])
                    suggested_prefixes.remove(suggested_prefixes[0])
                clear()
                continue
            elif p in prefixes:
                if p in orig_suggested_prefixes:
                    suggested_prefixes.append(p)
                    suggested_prefixes = sorted(list(suggested_prefixes), key=lambda x: (len(x), x))
                prefixes.remove(p)
                clear()
                continue
            elif p not in prefixes:
                try:
                    suggested_prefixes.remove(p)
                except:
                    pass
                prefixes.append(p)
                clear()
                continue
        except IndexError:
            clear()
            continue
    
    # generates all possible swaps of prefixes
    def generate_swapped_strings(input_str, characters_to_swap):
        def swap_characters(char):
            if char in characters_to_swap:
                return characters_to_swap[char]
            else:
                return [char]

        combinations = [input_str]

        for i, char in enumerate(input_str):
            if char in characters_to_swap:
                new_combinations = []
                for replacement in swap_characters(char):
                    for combination in combinations:
                        new_combination = combination[:i] + replacement + combination[i+1:]
                        new_combinations.append(new_combination)
                combinations.extend(new_combinations)

        return combinations
    
    characters_to_swap = {"I": ["1", "l", "i"], "1": ["I", "l", "i"], "l": ["I", "1", "i"], "i": ["I", "1", "l"], "O": ["D", "0", "Q"], "D": ["O", "0", "Q"], "0": ["D", "O", "Q"], "Q": ["D", "O", "0"], "5": ["S"], "S": ["5"]}

    # if character conversion errors are chosen, prefix list is updated with all possible combinations of each original prefix
    new_list = []
    if widen_boolean:
        for prefix in prefixes:
            combos = generate_swapped_strings(prefix, characters_to_swap)
            for combo in combos:
                new_list.append(combo)

        prefixes = new_list

    return prefixes

def create_regex():
    # creating regex patterns from prefixes and options

    if whitespace == ' ':
        flexr = '.?' * flex
        space = '.?'
    else:
        flexr = ''
        space = ''

    neg_lookaheads = []
    current_nla = []
    for letter in letters:
        for prefix in prefixes:
            if prefix[0] == letter:
                current_nla.append(prefix[1:])
        if current_nla == []:
            neg_lookaheads.append('[\._-]{0,2}?(?:' + whitespace + '){0,' + str(flex) + '}' + letter)
        else:
            neg_lookaheads.append('[\._-]{0,2}?(?:' + whitespace + '){0,' + str(flex) + '}' + letter + '(?!' + '|'.join(current_nla) + ')')

    neg_lookahead = '(?:' + '|'.join(neg_lookaheads) + ')?'

    patterns = []
    for prefix in prefixes:
        pattern = flexr.join(prefix) + space + '\d{0,1}?' + space + '(?:[\._-]{0,2}?(?:' + whitespace + '){0,' + str(flex) + '}' + widen + suffix
        patterns.append(pattern.replace('negative-lookahead-swap', neg_lookahead))

    # combines regex patterns into 1 to allow for order to be maintained in parsing
    return '|'.join(patterns)

def prefix_extract(match):
        # function to extract prefix from a given doc id

        limit = len(match) - 1
        prefix = ''
        index = -1
        for c in match:
            index += 1
            if c.isnumeric():
                if limit < index + 1:
                    break
                if not match[index + 1].isnumeric():
                    prefix += c
                break
            elif c in [' ', '-', '_', '.']:
                if limit < index + 1:
                    break
                if match[index + 1].isalpha():
                    prefix += c
            else:
                prefix += c

        return prefix

def suffixing_yes_no(match):
    if any(match.endswith(letter) for letter in letters):
        return 'Yes'
    else:
        return 'No'

def search():
    # conducting searching of text using regex patterns

    # choose whether to de-dupe Ids
    clear()
    while True:
        choices = {'1': True, '2': False}
        dedupe_choice = input('\n    De-duplicate doc IDs?  (On document level)\n\n    1 = Yes,  2 = No:  ')
        
        if dedupe_choice in choices:
            dedupe = choices[dedupe_choice]
            break
        else:
            clear()
            print('\n    Choice not valid.')
            continue

    if sys.getsizeof(files) > 150 or len(patterns) > 1:
        clear()
        print('\n    Generating results ...')

    hit_count = 0
    non_dedupe_results = []

    # swaps document text for a list of regex matches for each document
    for file in files:
        matches = []
        try:
            finds = re.findall(patterns, file[1])
            for find in finds:
                matches.append(find.strip())
        except:
            pass

        for match in matches:
            non_dedupe_results.append([match, file[0], file[2]])

        if dedupe:
            hit_count += len(matches)
            deduplicated_list = []
            for item in matches:
                if item not in deduplicated_list:
                    deduplicated_list.append(item)
            file[1] = deduplicated_list
        else:
            file[1] = matches

    # temporary result list
    results_temp = []
    for file in files:
        for match in file[1]:
            results_temp.append([match, file[0], file[2], prefix_extract(match), suffixing_yes_no(match)])

    original_folder = []
    sub_folder = []
    sub_folder_exists = False
    prefix_column = 3

    for result in results_temp:
        if result[2] == '':
            original_folder.append(result)
        else:
            sub_folder_exists = True
            sub_folder.append(result)

    if not sub_folder_exists:
        for result in original_folder:
            result.pop(2)
        prefix_column = 2

    # sorts documents with no subfolder to above
    results = original_folder + sub_folder

    # adds Prefix Count column for each row
    prefix_counter = dict(Counter([result[prefix_column] for result in results]))
    for result in results:
        for key, value in prefix_counter.items():
            if result[prefix_column] == key:
                result.insert(-1, value)
                break
        else:
            result.insert(-1, 'Error')

    return results, dedupe, hit_count, sub_folder_exists, non_dedupe_results

def hit_report_generate():
    # generates hit report with prefix counts, both regular and de-duped counts (if de-dupe is chosen)

    header = []
    dedupe_prefix_results = []
    for result in results:
        header.append(prefix_extract(result[0]))
        dedupe_prefix_results.append([prefix_extract(result[0]), result[1]])
    header = list(set(header))

    non_dedupe_prefix_results = []
    for result in non_dedupe_results:
        non_dedupe_prefix_results.append([prefix_extract(result[0]), result[1]])

    files_matched = int(len(set([r[1] for r in results])))

    directories = []
    for result in results:
        if result[2] != '':
            directories = {value: key for key, value in [list(t) for t in set(tuple(lst) for lst in [[r[2], r[1]] for r in results])]}
            break

    # non dedupe block
    non_dedupe_results_dict = {}
    for result in non_dedupe_prefix_results:
        value = result[0]
        key = result[1]

        if key not in non_dedupe_results_dict:
            non_dedupe_results_dict[key] = [[item, 0] for item in header]

        non_dedupe_results_dict[key][next(i for i, sublist in enumerate(non_dedupe_results_dict[key]) if sublist[0] == value)][1] += 1

    non_dedupe_block = [[item for item in [f"'-- NON DE-DUPE PREFIX HITS{' (With Lines Combined) --' if combine else ''}", "-- Directory --" if directories != [] else [], *header] if item != []], *[[key] + [item[1] for item in value] for key, value in non_dedupe_results_dict.items()]]

    if dedupe:

        # dedupe block
        dedupe_results_dict = {}
        for result in dedupe_prefix_results:
            value = result[0]
            key = result[1]

            if key not in dedupe_results_dict:
                dedupe_results_dict[key] = [[item, 0] for item in header]

            dedupe_results_dict[key][next(i for i, sublist in enumerate(dedupe_results_dict[key]) if sublist[0] == value)][1] += 1

        dedupe_block = [[item for item in [f"'-- DE-DUPE PREFIX HITS{' (With Lines Combined) --' if combine else ''}", "-- Directory --" if directories != [] else [], *header] if item != []], *[[key] + [item[1] for item in value] for key, value in dedupe_results_dict.items()]]

        hit_report = dedupe_block + [[]] + [[]] + non_dedupe_block

    else:

        hit_report = non_dedupe_block

    if directories != []:
        for line in hit_report[1:files_matched + 1] + hit_report[files_matched + 4:]:
            line.insert(1, directories[line[0]])

    # additional formatting/sorting
    starting_index = 1 if not directories else 2
    ending_index = starting_index+len(hit_report[len(hit_report)-1])-starting_index
    sums = [[i, sum(line[i] for line in hit_report[1:files_matched + 1] + hit_report[files_matched + 4:])] for i in range(starting_index, ending_index)]
    order = [[item[0], rank + starting_index] for rank, item in enumerate(sorted(sums, key=lambda x: x[1], reverse=True))]
    ordered_hit_report = [[] if i == [] else [i[h] for h in range(starting_index)] + [i[x[0]] for x in order] for i in hit_report]
    hit_report = ordered_hit_report

    return hit_report

def save():
    # saving results

    if os.path.isdir(path):
        directory = path
    else:
        directory = os.path.dirname(path)
    os.chdir(directory)

    # function to generate next available file names
    def file_name_generate(name):
        files_uncombined = os.listdir(directory)
        files = ' '.join(os.listdir(directory))
        folders = re.findall(f'{name} \d+', files)
        versions = [int(result.split(' ')[1]) for result in folders]
        if versions == []:
            if f'{name}.csv' not in files_uncombined:
                file_name = f'{name}.csv' 
            else:
                file_name = f'{name} 2.csv'
        else:
            file_name = f'{name} {str(max(versions) + 1)}.csv'

        return file_name

    results_name = file_name_generate('Results')
    hit_report_name = file_name_generate('Hit_Report')
    
    # saving hit report
    with open(hit_report_name, 'w', newline='', encoding='utf-8') as output:
        writer = csv.writer(output)
        writer.writerows(hit_report)

    # saving results
    with open(results_name, 'w', newline='', encoding='utf-8') as output:
        writer = csv.writer(output)
        if sub_folder_exists:
                writer.writerow(['Reference', 'Source Document', 'Directory', 'Prefix', 'Prefix Count', 'Suffixing Letter'])
        else:
                writer.writerow(['Reference', 'Source Document', 'Prefix', 'Prefix Count', 'Suffixing Letter'])
        writer.writerows(results)

    results_path = os.path.join(directory, results_name)

    return hit_report_name, results_path, results_name

if __name__ == "__main__":
    try:
        files, ext, path = locate_files()
        combine, suffix, widen, whitespace, flex, letters, widen_boolean = options()
        files, non_combine = read_text(files)
        suggested_prefixes = suggest()
        prefixes = get_prefixes(suggested_prefixes)
        patterns = create_regex()
        results, dedupe, hit_count, sub_folder_exists, non_dedupe_results = search()
        hit_report = hit_report_generate()

        if len(results) > 0:
            hit_report_name, results_path, results_name = save()
            if dedupe:
                if hit_count == len(results):
                    clear()
                    if str(input(f"\n    {len(results)} doc ID(s) found.  There were no duplicate references.\n\n    See '{results_name}' and '{hit_report_name}'.\n\n    Hit enter to exit (1 to see results) ... ")) == '1':
                            os.system(f'open "{results_path}"')
                else:
                    clear()
                    if str(input(f"\n    {hit_count} search hit(s), de-duplicated to {len(results)} unique doc ID(s).\n\n    See '{results_name}' and '{hit_report_name}'.\n\n    Hit enter to exit (1 to see results) ... ")) == '1':
                        os.system(f'open "{results_path}"')
            else:
                clear()
                if str(input(f"\n    {len(results)} doc ID(s) found.\n\n    See '{results_name}' and '{hit_report_name}'.\n\n    Hit enter to exit (1 to see results) ... ")) == '1':
                    os.system(f'open "{results_path}"')
        else:
            clear()
            input("\n    No ID's found.  Hit enter to exit ...")
    except Exception as e:
        clear()
        print(f"\n    Unexpected error: {e}")
        input('\n    Press enter to exit ...')