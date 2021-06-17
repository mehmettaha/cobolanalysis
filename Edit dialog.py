# coding: utf-8

# In[10]:

from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import pygraphviz as pgv 
from graphviz import Digraph
from graphviz import Source


def disp(G, root):
    def indisp():
        path = filedialog.asksaveasfilename(parent = root)
        if path.endswith(".svg"):
            path = path[:-4]
        if path.endswith(".map"):
            G.draw(path, format = "cmapx")
        else:
            with open((path + "-source.txt"),"w") as sfile:
                sfile.write(G.string())
            G.draw((path + ".svg"), format = "svg")
    return indisp

def search(radio, root, graph):
    def insea():
        obje = ""
        val = root.searchbox.get()
        mode = radio.get()
        
        if mode == "node":
            try:
                obje = graph.get_node(val)
            except KeyError:
                pass
            if obje != "":
                ne_editor(obje, mode)
        elif mode == "edge":
            val2 = root.edge2.get()
            try:
                obje = graph.get_edge(val,val2)
            except KeyError:
                pass
            if obje != "":
                ne_editor(obje,mode )
        elif mode == "subgraph":
            obje = graph.get_subgraph("cluster_" + val)
            if obje is None:
                obje = graph.get_subgraph(val)
            if obje is not None:
                ne_editor(obje,mode)
        elif mode == "all":
            remove_box()
            radio.set("node")
            obje = graph
            ne_editor(obje, mode)
        if obje is None or type(obje) == str:
            messagebox.showerror("Could not find", f"No such {mode} was found.", parent = root)
    return insea

def subgraph_editor(obje):
    child = Tk()
    nodg = StringVar(child)
    Radiobutton(child, text = "Node attributes", variable = nodg, value = "nodeattr").pack()
    Radiobutton(child, text = "Edge attributes", variable = nodg, value = "edgeattr").pack()
    Radiobutton(child, text = "Graph attributes", variable = nodg, value = "graphattr").pack()


def ne_editor(obje, mode):
    row = 0
    column = 0
    child = Tk()
    entries = []
    if mode == "node":
        list_of_attributes = [("shape","Shape of the node"),("URL","URL on click(start with http:\\ for online)"), ("color","Line color"),
                          ("comment","Hidden comment on the node"), ("fontcolor","text color"),("fontname", "goes to font-family in html"),("fontsize", "def:14"),
                         ("height","0.02:unlimited"),("width","0.02:unlimited"), ("label","Text to overwrite node name (you can use HTML here)"),
                              ("style","style name"),("fillcolor","color name"), ("tooltip","Text on hover")]
        analyze = dict(obje.attr)
    elif mode == "edge":
        list_of_attributes = [("style", "style of the line (hidden, dashed etc)"),("color", "line color"),
                              ("label","text on the edge"),("weight","0.0:1.0"), ("arrowhead","see arrowtype"),
                              ("comment","comments only visible here"),("constraint","true or false"),
                              ("edgeURL","url on the edge"),("edgetooltip","text on hover"),("labelURL","URL on the label"),
                             ("labelfontcolor","head/tail label only"),("labelfontsize","head/tail only, default=14"),("headlabel","head label"),("taillabel", "tail label")]
        analyze = dict(obje.attr)
    elif mode == "subgraph":
        list_of_attributes = [("style", "solid, filled, striped etc"), ("fillcolor", "only works with some styles"),("label", "do not lose the original name"),
                              ("pencolor","outer frame color"),("rank","same,min,source,max,sink"),("lheight","he")]    
        analyze = dict(obje.graph_attr)
    
    elif mode == "all":
        analyze = dict(obje.graph_attr)
        
        list_of_attributes = [("bgcolor","background"),("color","color of cluster lines"),("nodesep","distance between nodes"),("ranksep","distance between ranks"),
                              ("overlap","allow overlap, bool"),("clusterrank","local for clusters, global for no clustering"),
                             ("concentrate","merge edges when available, bool"), ("newrank","experimental, leave empty to turn off"),("overlap","refer to:https://www.graphviz.org/doc/info/attrs.html#d:overlap"),("splines","false,curved,polyline,ortho,true,none"),
                             ("sep","separation between nodes"),("scale","scale the graph size"),("ratio","attempt to adjust ratio, compress")]
        
        
    for key, label in list_of_attributes:
        Label(child,text = key).grid(row = row, column = column)
        e = Entry(child)
        if key in analyze.keys():
            e.insert(END,analyze[key])
            
        e.grid(row = row, column = column + 1)
        
        entries.append((key,e))
        Label(child, text = label).grid(row = row, column = column +2)
        row = row + 1
    Button(child, text = "Quit", command = child.destroy).grid(row = row, column = 0)
    Button(child, text = "Save", command = save_changes(entries, obje, mode, child)).grid(row = row, column = 1)
    
def save_changes(entries,obje, mode, root):
    def insave():
        newattr = dict()
        for key, entry in entries:
            newattr[key] = entry.get()
        
        if mode in ["node","edge"]:
            for key, value in newattr.items():               
                obje.attr[key] = newattr[key]
            if obje.attr["label"] == "" and mode == "node":
                obje.attr["label"] = str(obje) #Reset to node name if empty
                
        elif mode == "subgraph":
            obje.graph_attr.update(**newattr)
            if obje.graph_attr["label"] == "":
                obje.graph_attr["label"] = str(obje) #Reset to node name if empty
        elif mode == "all":
            
            obje.graph_attr.update(**newattr)

        messagebox.showinfo("Saved successfully", "Changes saved to current graph (Use 'Save to SVG' to save to file.)", parent = root)
        root.destroy()
        
    return insave

def load_source():
    global main
    #graph_source_path = r"C:\Users\hcs195021\Desktop\Cobol original\deneme-source.txt"
    graph_source_path = filedialog.askopenfilename(parent = main, filetypes = [("dot source files","*.txt")])
    if graph_source_path == "":
        return
    if to_edit(graph_source_path):
        main.destroy()
    
def gv_viewer(pgv_graph, enginevar):
    def innerview():
        gr = Source(pgv_graph.string(),format = "svg", engine = enginevar.get())
        gr.view()
    return innerview

def add_box():
    global main2
    main2.edge2 = Entry(main2)
    main2.edge2.grid(row = 5, column = 2)

def remove_box():
    global main2
    try:
        main2.edge2.destroy()
    except AttributeError:
        pass
    
def to_edit(path):
    global main, main2
    try:    
        G = pgv.AGraph(path)
    except:
        messagebox.showerror("Error loading source code","Error while drawing the graph. Please confirm you chose a .txt file with a valid dot source code", parent = main )
        return False
    main2 = Tk()
    G.layout("dot")
    radio = StringVar(main2)
    radio.set("node")
    t = Label(main2, text = "Search for node/edge/class").grid(row = 0)
    
    n = Radiobutton(main2, text = "node", variable =radio, value = "node", command = remove_box)
    n.grid(row = 1)
    e = Radiobutton(main2, text = "edge", variable = radio, value = "edge", command = add_box)
    e.grid(row = 2)
    s = Radiobutton(main2, text = "subgraph", variable = radio, value = "subgraph", command = remove_box)
    s.grid(row = 3)
    s = Radiobutton(main2, text = "graph", variable = radio, value = "all", indicatoron = 0, command = search(radio, main2, G))
    s.grid(row = 4)
    
    main2.searchbox = Entry(main2)
    main2.searchbox.grid(row = 5)

    but = Button(main2, text = "Search", command = search(radio, main2, G)).grid(row = 6, column = 0)

    but2 = Button(main2, text = "Save to SVG", command = disp(G, main2)).grid(row = 7, column = 0)
    
    row = 7
    
    engines = ["dot", "neato", "fdp", "sfdp", "twopi", "circo"]
    engine = StringVar(main2)
    Label(main2, text = "Choose an engine:").grid()
    for i, enginename in enumerate(engines):
        row = row + 1
        Radiobutton(main2, text = enginename, value = enginename, variable = engine).grid()
        print(enginename)
    engine.set("dot")
    
    but3 = Button(main2, text = "View without saving", command = gv_viewer(G, engine)).grid()   
    
    return True

main = Tk()

main2 = ""

Label(main, text = "Welcome to graph editor! Choose your graph using the button below:").pack()
Button(main, text = "Load source file (as .txt dot source code)", command = load_source).pack()
  
main.mainloop()


# In[20]:

