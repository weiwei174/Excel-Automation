import fitz
from dictionary import old_to_new, sheet_names, series

pdf_names = sheet_names

def replace_id(content):
        print('content: ', content)
        for old_id, new_id in old_to_new.items():

            if f'{old_id} (FT)' in content:
                # added to avoid re-replacing fixed IDs, but cannot break because multiple IDs are contained in the same textbox
                if f'/ {old_id} (FT)' in content or f'/{old_id} (FT)':
                     continue
                content = content.replace(f'{old_id} (FT)', new_id)
                # print('FT: ', content)
                continue
            
            elif old_id in content:
                if f'{old_id} /' in content or  f'/ {old_id}' in content or f'{old_id}/' in content or  f'/{old_id}' in content:
                     continue
                content = content.replace(old_id, new_id)
                # print('pt: ', content)

        return content


for name in pdf_names:
    pdf = fitz.open(f'{series} PDFS\{name}-MET - WIP.pdf')
    print(name)

    for page in pdf:
        comments = page.annots()
        for comment in comments:
            info = comment.info
            content = comment.get_textbox(comment)
            if 'WWID' not in content:
                continue 
            
            new_comment = replace_id(content)
            if new_comment == content:
                 continue
            info['content'] = new_comment
            comment.set_info(info)
            comment.update()
            # print(comment.info['content'])

    pdf.save(f'NEW {series} PDFS\{name}-MET - WIP.pdf')
