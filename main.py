import docx2txt
import re
import codecs

def make_index_tbl(path):
    """
        Makes a List of all indexes in the paths.
        :param path: a path to each part
        :type str.
        :return: List of all indexes.
        :rtype: str[].
        """
    x = docx2txt.process(path)
    tbl = x.splitlines()
    index_name_tbl = []
    for val in tbl:
        if len(val) > 2 and val.find('רשימת') == -1 and not val.isdigit() and val.find("whatdid") == -1 \
                and val.find("references") == -1 and val.find("value.writer") == -1 and val.find("newValue") == -1:
            if val.count(',') < 3:
                t = val.strip()
                if t.endswith(','):
                    t = t[:-1]
                index_name_tbl.append(t)
            else:
                val = val.replace('\n', '')
                val = val.split(',')
                while len(val) > 2:
                    t = val[0] + ',' + val[1]
                    val.remove(val[0])
                    val.remove(val[0])
                    t = t.strip()
                    for z in t:
                        if z.isdigit():
                            t = t.replace(z, '')
                    t = t.strip()
                    if not index_name_tbl.__contains__(t):
                        index_name_tbl.append(t)
    list(filter((2).__ne__, index_name_tbl))
    return index_name_tbl

# gets the path to the values and table of all index's, returns a table of Values.
def make_values_tbl(path, index_tbl):
    """
        Makes a List of all values in the paths, each member of the list must contain a name
        and one or more of the optionals: city, role, period.
        :param path: a path to each part
        :type str.
        :param index_tbl: List of all indexes.
        :type str[].
        :return: List of all values.
        :rtype: str[][4].
        """
    raw_text = docx2txt.process(path)
    split_text = raw_text.splitlines()
    values_tbl = []
    for i in range(0, len(split_text)):
        if index_tbl.__contains__(split_text[i]):
            if 0 < split_text[i + 2].count(",") < 4 and split_text[i + 2].count("ראה") == 0:
                x = split_text[i + 2].strip()
                tmp = re.split('; |,', x)
                if len(tmp) == 3:
                    if tmp[2].__contains__("מאח"):
                        tmp[2] = tmp[2].replace("מאח", "מאה")
                    p = (split_text[i], tmp[0], tmp[1], tmp[2])
                elif len(tmp) == 2:
                    if tmp[1].count("מאה") or tmp[1].count("מאח"):
                        if tmp[1].__contains__("מאח"):
                            tmp[1] = tmp[1].replace("מאח", "מאה")
                        p = (split_text[i], tmp[0], "---", tmp[1])
                    else:
                        p = (split_text[i], tmp[0], tmp[1], "---")
                elif len(tmp) == 1:
                    if tmp[0].count("מאה") > 0 or tmp[0].count("מאח") > 0:
                        if tmp[0].__contains__("מאח"):
                            tmp[0] = tmp[0].replace("מאח", "מאה")
                        p = (split_text[i], "---", "---", tmp[0])
                    else:
                        p = (split_text[i], tmp[0], "---", "---")
                else:
                    p = (split_text[i], "---", "---", "---")
            else:
                p = (split_text[i], "---", "---", "---")
            index_tbl.remove(split_text[i])
            values_tbl.append(p)
    return values_tbl


# parsing the index of both parts.
part1_index_path = 'C:\\Users\\moavl\\PycharmProjects\\DigitalHumanitiesMini\\part1_index.docx'
part2_index_path = 'C:\\Users\\moavl\\PycharmProjects\\DigitalHumanitiesMini\\part2_index.docx'
part1_index_tbl = make_index_tbl(part1_index_path)
part2_index_tbl = make_index_tbl(part2_index_path)
# parsing the values of both parts, using their index.
# after making the values table, the index table will get the index WITHOUT a value.
part1_values_path = 'C:\\Users\\moavl\\PycharmProjects\\DigitalHumanitiesMini\\part1_values.docx'
part2_values_path = 'C:\\Users\\moavl\\PycharmProjects\\DigitalHumanitiesMini\\part2_values.docx'
part1_values_tbl = make_values_tbl(part1_values_path, part1_index_tbl)
part2_values_tbl = make_values_tbl(part2_values_path, part2_index_tbl)
headers = ["שם", "תפקיד", "עיר", "תקופה"]
f = codecs.open("Persons Info - part1.txt", 'w', 'utf-8')
f.write("שם\tתפקיד\tעיר\tתקופה\n")
for x in part1_values_tbl:
    f.write(x[0] + '\t' + x[1] + '\t' + x[2] + '\t' + x[3] + '\t' + '\n')
f.close()
f = codecs.open("Persons Info - part2.txt", 'w', 'utf-8')
f.write("שם\tתפקיד\tעיר\tתקופה\n")
for x in part2_values_tbl:
    f.write(x[0] + '\t' + x[1] + '\t' + x[2] + '\t' + x[3] + '\t' + '\n')
f.close()

# opens a file which contains all indexes of part1 with no value, increasing lexicographically.
f = codecs.open("missing values - part1.txt", 'w', 'utf-8')
f.write("Total indexes: " + str((len(part1_values_tbl)) + (len(part1_index_tbl))) + '\n')
f.write("Number of Indexes with values: " + str((len(part1_values_tbl))) + '\n')
f.write("Number of Indexes with no values: " + str((len(part1_index_tbl))) + '\n')
for x in part1_index_tbl:
    f.write(x + '\n')
f.close()

# opens a file which contains all indexes of part2 with no value, increasing lexicographically.
f = codecs.open("missing values - part2.txt", 'w', 'utf-8')
f.write("Total indexes: " + str((len(part2_values_tbl)) + (len(part2_index_tbl))) + '\n')
f.write("Number of Indexes with values: " + str((len(part2_values_tbl))) + '\n')
f.write("Number of Indexes with no values: " + str((len(part2_index_tbl))) + '\n')
for x in part2_index_tbl:
    f.write(x + '\n')
f.close()


