from pptx import Presentation

prs = Presentation("C:/Users/j2017/OneDrive/School/pysc/Lectures_powerpoint/Class_2_Lecture_2_Part1.pptx")

# text_runs will be populated with a list of strings,
# one for each text run in presentation
text_runs = []

for slide in prs.slides:
    isFirst = True
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            if len(paragraph.text) > 0:
                if isFirst:
                    text_runs.append("* " + paragraph.text)
                    isFirst = False
                    print(str(len(paragraph.text)) + " " + "* " + paragraph.text)
                    break

                text_runs.append("** " + paragraph.text)
                print(str(len(paragraph.text)) + " " + "** " + paragraph.text)
    text_runs.append("\n")
