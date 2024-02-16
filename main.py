import c4d
from c4d import storage as s
import os
import csv


def render(rd, doc):
    bmp = c4d.bitmaps.MultipassBitmap(
        int(rd[c4d.RDATA_XRES]),
        int(rd[c4d.RDATA_YRES]),
        c4d.COLORMODE_RGB
    )
    if bmp is None:
        raise RuntimeError("Failed to create the bitmap.")

    bmp.AddChannel(True, True)

    if c4d.documents.RenderDocument(
        doc,
        rd.GetData(),
        bmp,
        c4d.RENDERFLAGS_EXTERNAL
    ) != c4d.RENDERRESULT_OK:
        raise RuntimeError("Failed to render the temporary document.")

    c4d.bitmaps.ShowBitmap(bmp)


def load_file(file, doc):
    c4d.documents.MergeDocument(
        doc,
        file,
        c4d.SCENEFILTER_OBJECTS | c4d.SCENEFILTER_MATERIALS
    )
    c4d.EventAdd()


def delete_file(doc):
    for obj in doc.GetObjects():
        if obj.GetName() == "Modell":
            obj.Remove()
    c4d.EventAdd()


def load_data(folder):
    data = []
    csv_file = [file for file in os.scandir(folder) if file.name == "times.csv"][0]
    with open(csv_file, "r") as file:
        reader = csv.reader(file)
        next(reader)
        for [name, start, end] in reader:
            data.append({"name": name, "start": start, "end": end})

    return data


def set_render(image, fps, rd, index):
    rd[c4d.RDATA_FRAMEFROM] = c4d.BaseTime(int(image["start"]), fps)
    rd[c4d.RDATA_FRAMETO] = c4d.BaseTime(int(image["end"]), fps)
    rd[c4d.RDATA_PATH] = os.path.join('out', str(index))
    c4d.EventAdd()


def delete_folder(dir):
    for file in os.scandir(dir):
        if file.is_file():
            os.remove(file)
        else:
            delete_folder(file)
    os.rmdir(dir)


def cleanup(folder):
    c4d.StatusSetText("Cleaning up...")
    files = os.scandir(folder)
    for file in files:
        if file.is_dir() and file.name.endswith(".fbm"):
           delete_folder(file)
    c4d.StatusClear()
    c4d.StatusSetText("Cleaning complete!")


def main():
    folder = s.LoadDialog(
        c4d.FILESELECTTYPE_ANYTHING,
        'Select folder to import',
        c4d.FILESELECT_DIRECTORY,
        ''
    )
    data = load_data(folder)
    doc = c4d.documents.GetActiveDocument()
    fps = doc.GetFps()
    rd = doc.GetActiveRenderData()
    color = c4d.Vector(1, 0, 0)
    c4d.StatusSetNetBar(0, color)
    for index, image in enumerate(data):
        c4d.StatusSetNetBar((index) / len(data) * 100, color)
        set_render(image, fps, rd, index)
        file = os.path.join(folder, image["name"] + '.fbx')
        load_file(file, doc)
        c4d.StatusClear()
        c4d.StatusSetText("Rendering " + image["name"] + "...")
        render(rd, doc)
        c4d.StatusSetText(f"Rendering {image['name']} complete!")
        delete_file(doc)
        c4d.EventAdd()
    c4d.StatusClear()
    c4d.StatusNetClear()
    c4d.StatusSetText("Rendering complete!")
    cleanup(folder)


if __name__ == '__main__':
    main()