import pandas as pd
import re
import xml.etree.ElementTree as ET
from xml.dom.minidom import parseString

def create_dico(input): #crée un dico clé valeur : tag:tag
    df = pd.read_csv(input, encoding='unicode_escape', on_bad_lines='warn', sep=r'\s*;\s*', engine='python')
    tag_map = {}

    for col_name in df.columns:
        match = re.search(r'\[(.+?)\]', col_name)
        if match:
            tag = match.group(1)
            tag_map[tag] = tag

    return tag_map

def get_tag_columns(df, tag_map): #crée un dico clé valeur : nom colonne:tag
    tag_columns = {}

    for col in df.columns:
        match = re.search(r'\[(.+?)\]', col)
        if match:
            col_tag = match.group(1)
            if col_tag in tag_map:
                tag_columns[tag_map[col_tag]] = col

    return tag_columns

def generate_element(tag, config, df, tag_columns, index):
    element = ET.Element(tag)

    for key, value in config.items():

        if 'attribute' in value and 'tag' not in value and len(value) == 1: #Exclue les balises qui ont des balises imbriquées
            sub_element = ET.Element(key)
            for attr_key, attr_value in value['attribute'].items():
                if attr_value in tag_columns:
                    sub_element.set(attr_key, str(df.at[index, tag_columns[attr_value]]))
            # Ajoute la balise avec uniquement des attributs
            element.append(sub_element)

        elif isinstance(value, dict):
            # Vérifie si la balise a un tag et des attributs
            if 'tag' in value and 'attribute' in value:
                sub_element = ET.SubElement(element, key)
                if value['tag'] in tag_columns:
                    sub_element.text = str(df.at[index, tag_columns[value['tag']]])
                for attr_key, attr_value in value['attribute'].items():
                    if attr_value in tag_columns:
                        sub_element.set(attr_key, str(df.at[index, tag_columns[attr_value]]))
                # Vérifie si la balise n'est pas vide
                if not sub_element.attrib and (sub_element.text is None or not sub_element.text.strip()):
                    element.remove(sub_element)

            # Gère les balises avec des boucles 'loop'
            elif 'loop' in value:
                    # Traiter la boucle à partir de la prochaine balise
                    if value['loop'] in tag_columns:
                        for i in df.index:
                            loop_element = generate_element(key, value, df, tag_columns, i)
                            # Traiter les attributs de la balise si présents
                            if 'attribute' in value:
                                for attr_key, attr_value in value['attribute'].items():
                                    loop_element.set(attr_key, str(df.at[index, tag_columns[attr_value]]))
                            if loop_element is not None and len(loop_element):
                                element.append(loop_element)

            # Gère les balises imbriquées
            else:
                # Vérifie si la clé est 'attribute' pour éviter la récursivité
                if key != 'attribute':
                    sub_element = generate_element(key, value, df, tag_columns, index)
                    # Vérifie si la balise n'est pas vide
                    if sub_element is not None: 
                        # Ouvre la balise avant l'exploration des balises imbriquées
                        if 'attribute' in value and 'tag' not in value:
                            for attr_key, attr_value in value['attribute'].items():
                                if attr_value in tag_columns:
                                    sub_element.set(attr_key, str(df.at[index, tag_columns[attr_value]]))
                        element.append(sub_element)

        #Gère les balises ne possédant pas d'attribut, ni d'imbriquation
        elif isinstance(value, str):
            if value in tag_columns and not pd.isna(df.at[index, tag_columns[value]]) and (key != 'loop') and (key != 'loopId'):
                sub_element = ET.SubElement(element, key)
                sub_element.text = str(df.at[index, tag_columns[value]])
                # Vérifie si la balise n'est pas vide
                if not sub_element.attrib and (sub_element.text is None or not sub_element.text.strip()):
                    element.remove(sub_element)

    # Supprime les éléments vides
    if not element.attrib and not list(element) and (element.text is None or not element.text.strip()):
        return None

    return element

def replace_virgule(xml_file, keys_to_process):
    # Parse XML and modify text nodes for specific keys
    parser = ET.XMLParser(remove_blank_text=True)
    tree = ET.parse(xml_file, parser)
    root = tree.getroot()
    
    for element in root.iter():
        if element.tag in keys_to_process and element.text is not None and ',' in element.text:
            element.text = element.text.replace(',', '.')

def handle_loop_id(tag, config, df, tag_columns):
    root_element = ET.Element(tag)
    seen_ids = set()

    loop_id_key = config['loopId']
    for id_value, df_group in df.groupby(tag_columns[loop_id_key]):
        if id_value in seen_ids:
            continue
        seen_ids.add(id_value)

        sub_elements = generate_element(tag[:-1], config, df_group, tag_columns, index=df_group.index[0])
        if sub_elements is not None and len(sub_elements):
            for sub_element in sub_elements:
                root_element.append(sub_element)

    return root_element

def remove_empty_elements(element):
    # Supprime les éléments enfants vides récursivement
    for child in list(element):
        remove_empty_elements(child)
        if len(child) == 0 and not child.attrib and (child.text is None or not child.text.strip()):
            element.remove(child)

def replace_virgule(element, keys):
    if element.tag in keys and element.text is not None and ',' in element.text:
        element.text = element.text.replace(',', '.')
    
    for child in element:
        replace_virgule(child, keys)

def format_date(date_str):
    date_str = date_str.replace('/', '-')
    # Solution pour chaque format de date possible
    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        return f"{date_str}T00:00:00"
    elif re.match(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$', date_str):
        return date_str.replace(' ', 'T')
    elif re.match(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$', date_str):
        return f"{date_str.replace(' ', 'T')}:00"
    elif re.match(r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$', date_str):
        return date_str
    elif re.match(r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$', date_str):
        return f"{date_str}:00"
    else:
        return f"{date_str}T00:00:00"

# Fonction pour formater les dates dans l'arbre XML
def dates_in_element(element):
    date_regex = re.compile(r'Date', re.IGNORECASE)
    for child in list(element):
        if date_regex.search(child.tag):
            if child.text:
                child.text = format_date(child.text)
        dates_in_element(child)


def CSV_to_XML(input_file, output, config):
    if not input_file.endswith('.csv'):
        print("Expected a CSV file")
        return
    if not output.endswith('.xml'):
        print("Expected an XML file")
        return
    
    try:
        df = pd.read_csv(input_file, encoding='unicode_escape', on_bad_lines='warn', sep=r'\s*;\s*', engine='python')
    except FileNotFoundError:
        print("CSV file not found")
        return
    
    tag_map = create_dico(input_file)
    tag_columns = get_tag_columns(df, tag_map)
    
    root_tag = list(config.keys())[0]
    root_config = config[root_tag]
    root_element = ET.Element(root_tag)
    
    for key, value in root_config.items():
        if 'loopId' in value:
            loop_element = handle_loop_id(key, value, df, tag_columns)
            if len(loop_element):
                root_element.append(loop_element)
        else:
            element = generate_element(key, value, df, tag_columns, index=0)
            if element is not None and len(element):
                root_element.append(element)

    remove_empty_elements(root_element)

    replace_virgule(root_element, ['NumericalValue', 'ReceiverAlphaNumericalValue'])

    dates_in_element(root_element)
                  
    tree = ET.ElementTree(root_element)
    tree.write(output, encoding='utf-8', xml_declaration=True)
    
    with open(output, 'r', encoding='utf-8') as f:
        xml_str = f.read()
    dom = parseString(xml_str)
    pretty_xml = dom.toprettyxml()
    
    with open(output, 'w', encoding='utf-8') as f:
        f.write(pretty_xml)

    return pretty_xml

xml_config = {
    'SampleDocument': {
        'Samples': {
            'loopId': 'Id', #INDICATEUR BOUCLE
            'Sample': {
                'Id': 'Id',
                'IntakeCode': 'inn',
                'SamplingDate': 'dt_take',
                'SendDate': 'dt_send',
                'ProductionDate': 'dt_pd',
                'Product': {'tag': 'pd_nm', 'attribute': {'Type': 'pd_type', 'Id': 'pd_code'}},
                'LibraCode': {'tag': 'nr_la_nm', 'attribute': {'Id': 'nr_la_code'}},
                'VersionNumber': 've',
                'Supplier': {'tag': 'sup_nm', 'attribute': {'Id': 'sup_nr'}},
                'Origin': {'tag': 'orig_nm', 'attribute': {'Id': 'orig_nr'}},
                'ExternalComment': 'comm_ext',
                'SamplingPlan': {'tag': 'mop_nm', 'attribute': {'Id': 'mop_nr'}},
                'System': {'tag': 'sy_nm', 'attribute': {'Id': 'sy_code'}},
                'Customer': {'tag': 'cust_nm', 'attribute': {'Id': 'cust_code'}},
                'SalesRepresentative': {'tag': 'tech_nm', 'attribute': {'Id': 'tech_code'}},
                'BillingCustomer': {'tag': 'cust_nm_fact', 'attribute': {'Id': 'cust_code_fact'}},
                'SamplingLocation': {'tag': 'place_nm', 'attribute': {'Id': 'place_nr'}},
                'Class': {'tag': 'class_nm', 'attribute': {'Id': 'class_nr'}},
                'Grade': {'tag': 'grd_nm', 'attribute': {'Id': 'grd_nr'}},
                'SampledBy': 'sampled_by',
                'Farmer': 'farmer',
                'OrderCode': 'order_code',
                'Transport': 'tr',
                'InternalComment': 'comm_int',
                'ExternalComment': 'comm_ext',
                'CountryCode': 'country',
                'ZoneCode': 'zone',
                'PlaceCode': 'place',
                'Loader': {'tag': 'loader_nm', 'attribute': {'Id': 'loader_nr'}},
                'SampleType': {'tag': 'sample_type_nm', 'attribute': {'Id': 'sample_type_nr'}},
                'Results': {
                    'Result': {
                        'loop': 'Id', #INDICATEUR BOUCLE
                        'Criterion': {
                            'attribute': {'Id': 'nu_code'},
                            'Name': 'nu_nm',
                            'Method': {'tag': 'met_nm', 'attribute': {'Id': 'met_nr'}}
                        },
                        'Lab': {'tag': 'lab_nm', 'attribute': {'Id': 'lab_nr'}},
                        'NumericalValue': 're',
                        'ReceiverAlphaNumericalValue': 're',
                        'Unit': {'attribute': {'Id': 'uom'}}
                    }
                },
                'SampleLevels': {
                    'Level': {'attribute': {'Id': 'level_nr'}}
                }
            }
        }
    }
}

xml_config_test= {
    'Document': {
        'Header' : {
            'loopId':'Id',
            'TypeObj': {'tag':'type', 'attribute': {'Id': 'code', 'Type':'code2'}},
            'Texte': {'tag':'txt', 'attribute': {'Id' : 'txt_code'}}
        },
        'Body' : {
            'loopId':'Id',
            'ID':'Id',
            'Object': {
                'attribute': {'Name':'obj'},
                'loop':'Id',
                'ObjectLvl': {
                    'attribute':{'Id' : 'lvl'},
                    'Version':'ve',
                    'Value': {'tag':'val', 'attribute':{'Id': 'val_code'}}
                },
                'loop':'Id',
                'ObjectLength':{
                    'attribute':{'Len':'len'},
                    'Ratio':'ra'
                }
            }
        }
    }
}

CSV_to_XML('csv_sample_tag.csv', 'testFinal2.xml', xml_config)
#CSV_to_XML('test_csv2.csv', 'testFinal3.xml', xml_config_test)