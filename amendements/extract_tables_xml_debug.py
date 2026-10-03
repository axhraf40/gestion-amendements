import zipfile
from lxml import etree

# Fonction récursive pour parcourir le XML et afficher les tableaux imbriqués
def print_tables_xml(element, level=0):
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    for child in element:
        if child.tag.endswith('tbl'):
            print(f"{'  '*level}Tableau (niveau {level}) trouvé dans le XML.")
            print_tables_xml(child, level+1)
        else:
            print_tables_xml(child, level)

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python extract_tables_xml_debug.py chemin/vers/fichier.docx")
        sys.exit(1)
    docx_path = sys.argv[1]
    # Ouvrir le fichier docx comme une archive zip
    with zipfile.ZipFile(docx_path) as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            tree = etree.parse(xml_file)
            root = tree.getroot()
            print("Analyse des tableaux dans le XML :")
            print_tables_xml(root) 