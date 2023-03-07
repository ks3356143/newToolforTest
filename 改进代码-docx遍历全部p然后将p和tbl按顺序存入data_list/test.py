from docx import Document

doc = Document('test.docx')

data_list = []

for ele in doc._element.body.iter():
    data = {'type':''}
    if ele.tag.endswith('}p'):
        elePstyle = ele.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
        if len(elePstyle) >= 1:
            # 获取标题级别
            rank = elePstyle[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            t_ele = ele.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            title = '' 
            for i in range(len(t_ele)):
                title = title + t_ele[i].text
            data['type'] = rank
            data['title'] = title
            data_list.append(data)
    elif ele.tag.endswith('}tbl'):
        data['type'] = 'table'
        data['other'] = '是表格'
        data_list.append(data)
print(data_list)

