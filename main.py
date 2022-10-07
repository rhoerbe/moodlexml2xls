import pandas
from lxml import etree
from pathlib import Path
import html2text

def main():
    basepath = Path(r"c:\Users\r2h2\Downloads")
    fp_moodle_questionbank_xml = basepath / "quiz-mis22w22is-iam-Standard für mis22w22is-iam-20220929-0846.xml"
    fp_moodle_questionbank_xls = basepath / "quiz-mis22w22is-iam-Standard für mis22w22is-iam-20220929-0846.xlsx"
    full_tree = parse_xml(fp_moodle_questionbank_xml)
    question_tree = get_questions(full_tree)
    taglist = get_question_tags(full_tree)
    d = create_table_dict(question_tree, taglist)
    d = map_tree_to_dict(question_tree, taglist, d)
    write_tableformat(d, fp_moodle_questionbank_xls)


def parse_xml(p: Path) -> object:
    try:
        tree = etree.parse(p)
    except UnicodeDecodeError as e:
        print(e)
        exit(1)
    return tree


def get_questions(full_tree) -> list:
    #question_tree = full_tree.xpath("/quiz/question[@type='multichoice']")
    question_tree = full_tree.xpath("/quiz/question[@type='essay' or @type='multichoice' or @type='shortanswer']")
    print(f"reading {len(question_tree)} multichoice questions")
    return question_tree


def get_question_tags(full_tree) -> set:
    # compiling set of used tags'
    taglist = set()
    for tag in full_tree.xpath("/quiz/question[@type='essay' or @type='multichoice' or @type='shortanswer']/tags/tag/text"):
        taglist.add(tag.text.strip())
    return taglist


def create_table_dict(question_tree, taglist) -> dict:
    # to create an excel table we make a dict of columns (col headers are dict keys)
    d = dict()
    d['name'] = []
    d['type'] = []
    d['questiontext'] = []
    d['grade'] = []
    for t in sorted(taglist):
        d[t] = []
    return d


def map_tree_to_dict(question_tree, taglist, d) -> dict:
    for q in question_tree:
        d['name'].append(q.xpath("./name/text")[0].text.strip())
        d['type'].append(q.xpath("./@type")[0].strip())
        question_html = q.xpath("./questiontext/text")[0].text.strip()
        d['questiontext'].append(html2text.html2text(question_html))
        d['grade'].append(int(float(q.xpath("./defaultgrade")[0].text.strip())))
        alltags = dict.fromkeys(taglist, '')
        usedtags = q.xpath("./tags/tag/text")
        for usedtag_elem in usedtags:
            usedtag = usedtag_elem.text.strip()
            alltags[usedtag] = 'X'
        for t in sorted(taglist):
            d[t].append(alltags[t])
    return d


def write_tableformat(d, out_path):
    df = pandas.DataFrame(d)
    print(df)
    df.to_excel(out_path, index=False)

if __name__ == '__main__':
    main()

