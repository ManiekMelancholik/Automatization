import SpssClient



def export_data():
    path = 'C:/Users/Quiqhaqru/Desktop/SpssTxtPython.txt'
    file = open(path, 'w')
    file.write("Export table from procedure :\n")

    SpssClient.StartClient()
    currentOut = SpssClient.GetDesignatedOutputDoc()
    currentOutItems = currentOut.GetOutputItems()

    for i in range(currentOutItems.Size()):
        outItem = currentOutItems.GetItemAt(i)
        if outItem.GetType() == SpssClient.OutputItemType.TEXT:
            tempItem = outItem.GetSpecificType()
            file.write(tempItem.GetTextContents())
            file.write(f"\n#@#: {i}\n\n")

    print("_____Done_____")
