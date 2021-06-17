#!/usr/bin/env python
# coding: utf-8

# In[1]:


from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
from graphviz import Digraph
import sys 
from math import sqrt


                     
class Metamodule:
    # Attributes:
    # graph : the Digraph object for display
    # __excel : the pd.ExcelFile object for reading from file
    # __cluster_umu(有無) : the value for whether the graph will be created with clusters
    # __class_base : The class table to make it easier to pass it between functions
    # __relation_base : The relations table to make it easier to pass it between functions
    # __nodelist : The complete list of nodes, and their classes
    # __rellist : The complete list of relations, and the classes of both instances

    """ 
    This class serves as a package for various functions of the Metamoduling tool. 
    On initialization, the class creates the source code for the graph, which enables the functions of the tool.
    You can access those functions with options() method, or run the methods directly.
    """
    
    def print_attributes(self):
        print("Clusters:" + self.__cluster_umu)
        print("Class base:")
        print(self.__class_base)
        print("Relation base:")
        print(self.__relation_base)
        print("Node list:")
        print(self.__nodelist)
        print("Relation list:")
        print(self.__rellist)

    def info(self):
        print(self.__doc__)

    def __init__(self, address, format="svg", engine="dot", ranksep="2", overlap="false", cluster_umu="yes", cluster_samerank = "no", weight = 0):
        try:
            self.__excel = pd.ExcelFile(address)
        except:
            raise ("Could not load file")
            
        self.graph = Digraph(engine=engine, format=format)
        self.graph.attr(ranksep=ranksep, overlap=overlap)
        self.__cluster_umu = cluster_umu
        self.run_checks()
        
        self.make_graph(cluster_umu, cluster_samerank, weight)
        
    def draw_edges(self):
        
        for _, frame in self.__rellist.reset_index().iterrows():
            self.graph.edge(frame["Instance1"],
                    frame["Instance2"], color = frame["Line Color"], style = frame["Line Type"], label = frame["Label"])

    def run_checks(self):
        checked = self.check_excel(self.__excel)
        if type(checked) == SyntaxError:
            raise checked
        try:
            self.__class_base = self.read_classes(self.__excel)
        except:
            raise ("Class table not found")

        try:
            self.__relation_base = self.read_relations(self.__excel)
        except:
            raise ("Relation table not found")

    def make_graph(self,cluster_umu, cluster_samerank, weight):
        
        self.__rellist = self.edging(
            self.graph, self.__relation_base, self.__excel)
        
        self.__rellist = pd.DataFrame(self.__rellist)
        if self.__rellist.empty == True:
            raise NoRelationException
        self.__rellist.columns = ["Instance1", "Instance2","Line Type", "Line Color", "Label"]
        self.__rellist = self.__rellist.set_index(["Instance1","Instance2"])
        
        self.__nodelist = self.instance_nodes(
            self.graph, self.__class_base, self.__excel, cluster_umu, cluster_samerank, weight)
        self.__nodelist = pd.DataFrame(self.__nodelist, columns=[
                                       "Instance Name", "Instance class"]).set_index("Instance Name")

        self.draw_edges()      
        
    def view(self):
        self.graph.view()

    def check_excel(self, excel):  # Excelのフォーマットをチェックする
        for i in excel.sheet_names:
            frame = list(pd.read_excel(excel, i).columns)  # ヘッダー名のリスト
            if i == "Class":
                if frame != ["No", "Class Name", "Shape", "Line Color", "Fill Color", "remarks"]:
                    return SyntaxError("Check 'Class' table headers")
            elif i == "Relation":
                if frame != ["No", "Class Name 1", "Class Name 2", "Line Type", "Line Color", "remarks", "Relation Type", "Direction"]:
                    return SyntaxError("Check 'Relation' table headers")
            elif i.startswith("relation"):
                for x in ["No", "ClassX instance", "ClassY instance", "Label"]:
                    if x not in frame:  # その他のヘッダーもある可能性があるため完全一致はもとめない
                        return SyntaxError("Check " + i)
            elif i.startswith("Instance_"):
                i = i[9:]
                for x in ["No", "Instance Name", "Shape", "Line Color", "Fill Color"]:
                    if x not in frame:
                        return SyntaxError("Check " + i)
        return 0

    def read_classes(self, excel):  # 元のExcelのクラス定義表を読み込む。クラス別にスタイルや色の設定に必要
        classes = pd.read_excel(excel, "Class")
        classes = classes.set_index("Class Name")
        classes = classes.fillna("")  # 無い項目を空にする
        return classes  # あとでインスタンス関数に渡す

    # 元のExcelの関係定義表を読み込む。クラスのペアで関係のやじるしのスタイルの設定に必要
    def read_relations(self, excel):
        relations = pd.read_excel(excel, "Relation")
        relations = relations.set_index("No")
        relations = relations.fillna("")  # 無い項目を空にする
        return relations  # あとで関係の関数に渡す

    def edging(self, g, relation_base, excel):  # 関係のやじるしの作成
        rellist = list()  # 全関係をまとめるlist
        for i in excel.sheet_names:  # Excelのシート名をスキャンする

            if i.startswith("relation"):  # 関係表は「relation」で始めなくてはいけない
                frame = pd.read_excel(excel, i).fillna("")
                # 関係定義表で、シート名に一致する行を見つけ出す
                index = relation_base[relation_base["remarks"] == i].iloc[0]

                with g.subgraph() as c:
                    # 関係定義表から属性を読み込む
                    c.attr(
                        "edge", style=index["Line Type"], color=index["Line Color"], dir=index["Direction"])
                    
                    for x in range(len(frame)):
                        color = frame["Line Color"][x]
                        style = frame["Line Type"][x]
                        label = frame["Label"][x]
                        if color == "":
                            color = index["Line Color"]
                        if style == "":
                            style = index["Line Type"]
                    
                        try:  # tryしないとClassXとClassYが入っていない行でもedgeになってしまう
                            #c.edge(frame["ClassX instance"][x],
                                 #  frame["ClassY instance"][x], color = color, style = style, label = label)
                            rellist.append(
                                [frame["ClassX instance"][x], frame["ClassY instance"][x],style,color,label])
                        except:
                            break
        return rellist

    def instance_nodes(self, g, class_base, excel, cluster_umu, cluster_samerank, weight):
        ccc= type(weight)
        print(f"weight = {weight} with type:{ccc}")
        nodes = list()
        for x in excel.sheet_names:
            if x.startswith("Instance_"):
                frame = pd.read_excel(excel, x)
                classname = x[9:]
                attrs = class_base.loc[classname]
                if cluster_umu == "yes":
                    subname = "cluster_" + classname
                elif cluster_umu == "no":
                    subname = classname
                else:
                    print("Invalid value for cluster, using 'yes'")
                    cluster = "yes"

                with g.subgraph(name=subname) as c:
                    c.attr(label=classname,  nodesep="2")
                    if cluster_samerank == "yes" or cluster_samerank == True:
                        c.attr(rank = "same")
                    for i in range(len(frame)):
                        c.attr(
                            "node", shape=attrs["Shape"], color=attrs["Line Color"], width="2")
                        if attrs["Fill Color"] != "":
                            c.attr("node", style="filled",
                                   fillcolor=attrs["Fill Color"])
                        self.check_and_update("Shape", i, frame, c)
                        self.check_and_update("Line Color", i, frame, c)
                        self.check_and_update("Fill Color", i, frame, c)
                        nodename = frame.iloc[i]["Instance Name"]
                        rle = self.__rellist.reset_index()
                        
                        toolforward = list()
                        for _, line in rle[rle["Instance1"] == nodename].iterrows():
                            toolforward.append(line["Instance2"])
                        toolbackward = list()
                        for _, line in rle[rle["Instance2"] == nodename].iterrows():
                            toolbackward.append(line["Instance1"])
                            
                        tooltip = f"Node ID: {nodename} &#010;Class: {classname} &#010;Affects: {toolforward} &#010;Affected by: {toolbackward}"
                        
                        if weight == 1:
                            
                            importance = len(toolforward) + len(toolbackward)

                            try:
                                c.node(nodename, tooltip = tooltip, width = str(importance**(2/3)), height = str(importance**(2/3)/2),
                                       fontsize = str(importance**(2/3)*12)) #graphviz needs str format not int
                                nodes.append([nodename, classname])
                            except: 
                                break
                            
                        elif weight == 0:
                            try:
                                c.node(nodename, tooltip = tooltip)
                                nodes.append([nodename, classname])
                            except:
                                break
        return nodes

    def check_and_update(self, colname, i, frame, c):
        if colname == "Shape":
            try:
                shape = frame.iloc[i][colname]
                if type(shape) == str:
                    c.attr("node", shape=shape)
            except:
                pass
        elif colname == "Line Color":
            try:
                linecolor = frame.iloc[i]["Line Color"]
                if type(linecolor) == str:
                    c.attr("node", color=linecolor)
            except:
                pass
        elif colname == "Fill Color":
            try:
                fillcolor = frame.iloc[i]["Fill Color"]
                if type(fillcolor) == str:
                    c.attr("node", fillcolor=fillcolor)
            except:
                pass

    def filter_att(self, excel, graph, nodelist, class_base):
        classchoice = input("Select the class you want to work with:\n")
        print("You chose" + classchoice)
        classchoice = "Instance_" + classchoice
        instable = pd.read_excel(excel, classchoice)
        print("Available attributes:")
        print(list(instable.columns))
        attchoice = input("Which attribute do you want to filter?:\n")
        val = input("Which value do you want to see?:")
        engine = input("Which engine?\n -dot -circo -sfdp -twopi")
        if engine == "":
            engine = "dot"

        with open("temp.txt", "w") as file:
            file.write(graph.source)
        g = Digraph(format="svg", engine=engine)
        all_alls = list()    #Alls records all relations for a single instance, which are then collected here 
        
        for _, row in instable.iterrows():
            try:
                if row[attchoice] == val:
                    alls = self.focus_instance(
                        row["Instance Name"], "temp.txt")
                    all_alls.append(alls)
            except:
                raise AttributeError("Can't find")
         # "temp.txt"
        g.attr(overlap="false")
        for alls in all_alls:
            for i in alls:
                for x in i:
                    try:

                        att = class_base.loc[nodelist.loc[x].values[0]]
                    except:
                        print("failed here:")
                        print(i, x)
                        raise ExceptionError

                    if type(att) == "pd.DataFrame":
                        print("Node" + x + "is in more than one class")
                        raise AttributeError

                    try:
                        g.node(x, shape=att["Shape"], style="filled",
                               color=att["Line Color"], fillcolor=att["Fill Color"])
                    except:
                        print("Failed at:")
                        print(x)
                        print("Trying without attributes:")
                        g.node(x)

        for alls in all_alls:
            for i in alls:
                g.edge(i[0], i[1])
        g.view()

    def wrap(self, first, last, rellist):  # recursiveの機能を使いやすくするためのwrapper関数
        collect = set()
        for g in first:
            head = g
            self.recursive(rellist, g, last, head, collect)
        return collect

    def recursive(self, frame, x, last, head, collect):
        mid = frame[frame["Instance1"] == x]
        if mid.empty is False:
            for i in mid.values:
                if i[1] == head:
                    return
                if i[1] in last:
                    collect.add((head, i[1]))
                else:
                    self.recursive(frame, i[1], last, head, collect)

    def classrelation(self, class1, class2, engine):
        fir = list()
        las = list()
        for i in self.__nodelist.iterrows():
            if class1 == i[1][0]:
                fir.append(i[0])
        if len(fir) == 0:
            message = "Cannot find class " + class1
            return message

        for i in self.__nodelist.iterrows():
            if class2 == i[1][0]:
                las.append(i[0])
        if len(las) == 0:
            message = "Cannot find class " + class2
            return message
        if engine == "":
            engine = "dot"
        h = Digraph(format="svg", engine=engine)
        rela = self.wrap(fir, las, self.__rellist.reset_index())
        if rela == set():
            rela = self.wrap(las, fir, self.__rellist.reset_index()) #try in reverse in case classes were entered in reverse
        with h.subgraph(name="cluster_first") as c:
            c.attr(label=class1)
            c.attr("node", color="red")
            for i in fir:
                c.node(i)
        with h.subgraph(name="cluster_second") as c:
            c.attr(label=class2)
            c.attr("node", color="blue")
            for x in range(len(las)):
                c.attr(rank="same")
                c.node(las[x])

        for i in rela:
            h.edge(i[0], i[1], minlen="3")
        return h

    def inputter(self, rellist, nodelist):
        class1 = input("Enter class 1: ")
        class2 = input("Enter class 2: ")
        print("Enter engine(circo by default):\ndot\nneato\ncirco\nfdp\ntwopi\nsfdp\n")
        engine = input()
        h = self.classrelation(class1, class2, rellist, nodelist, engine)
        if type(h) == Digraph:
            h.view()
        else:
            print(h)
            raise NoRelationFound("Nothing to draw")

    def class_diagram(self):
        g = Digraph(format="svg", strict="True")
        g.attr(rankdir = "BT")
        g.attr("edge", dir = "back")
        for i in self.__class_base.index:
            g.node(i)
        alist = list()
        blist = list()

        for i in range(len(self.__rellist)):
            a = self.__nodelist.loc[self.__rellist.iloc[i].name[0]]
            if a.empty is False:
                alist.append(a["Instance class"])
            else:
                alist.append("")
            b = self.__nodelist.loc[self.__rellist.iloc[i].name[1]]
            if b.empty is False:
                blist.append(b["Instance class"])
            else:
                blist.append("")
        
       
        for i in zip(alist,blist):
            try:
                g.edge(i[1], i[0])
            except:
                pass
        return g

    def instance_go(self, temp, alls, done, focus="", mode=""):
        flag = False
        for line in temp:
            toke = [i.strip('"\t\n') for i in line.split()][:3]
            if mode == "b":
                x = 2
            elif mode == "f":
                x = 0
            try:
                if focus == toke[x]:
                    alls.add((toke[0], toke[2]))
                    flag = True
                    done.add(focus)
            except (IndexError, KeyError):
                pass
        return flag

    def focus_instance(self, instance, source):
        alls1 = set()
        alls2 = set()
        done = set()
        temp = list()
        flag = 0

        for line in source.split("\n"):
            if flag == 0:
                if "->" in line:  # Starts reading after seeing "->" which means the first edge
                    flag = 1
            if flag == 1:
                if line.strip().startswith("edge") or line.strip().startswith("{") or line.strip().startswith("}"):
                    pass
                else:
                    temp.append(line)
        print(temp)
        ctr1 = 0
        ctr2 = 0

        while True:
            if ctr1 == 0:
                self.instance_go(temp, alls1, done, focus=instance, mode="b")
                ctr1 = ctr1 + 1
            else:
                ctr = len(done)
                for pair in list(alls1):
                    for x in pair:
                        if x not in done:
                            self.instance_go(
                                temp, alls1, done, focus=x, mode="b")
                if ctr == len(done):
                    break

        while True:
            if ctr2 == 0:
                self.instance_go(temp, alls2, done, focus=instance, mode="f")
                ctr2 = ctr2 + 1
            else:
                ctr = len(done)

                for pair in list(alls2):
                    for x in pair:
                        if x not in done:
                            self.instance_go(
                                temp, alls2, done, focus=x, mode="f")

                if ctr == len(done):
                    break

        alls = set()
        for i in list(alls1):
            alls.add(i)
        for i in list(alls2):
            alls.add(i)

        return alls

    def focus_input(self, instance, engine):
        g = Digraph(format="svg", engine=engine)
        alls = self.focus_instance(instance, self.graph.source)
        g.attr(overlap="false")

        for i in alls:
            for x in i:
                try:
                    att = self.__class_base.loc[self.__nodelist.loc[x].values[0]]
                except:
                    print("failed here:")
                    print(i, x)
                if type(att) == "pd.DataFrame":
                    print("Node" + x + "is in more than one class")

                try:
                    tooltip = self.__nodelist.loc[x]["Instance class"]
                    g.node(x, shape=att["Shape"], style="filled",
                           color=att["Line Color"], fillcolor=att["Fill Color"], tooltip = f"Class: {tooltip}")
                except:
                    print("Failed at:")
                    print(x)
                    print("Trying without attributes:")
                    g.node(x)
        
        for i in alls:
            direction = 0
            print(i)
            try:
                tab = self.__rellist.loc[i[0],i[1]]
            except:
                tab = self.__rellist.loc[i[1],i[0]]
                direction = 1
            for _, row in tab.iterrows():
                if direction == 0:
                    g.edge(i[0],i[1], style = row["Line Type"], color = row["Line Color"], label = row["Label"])
                else:
                    g.edge(i[1],i[0], style = row["Line Type"], color = row["Line Color"], label = row["Label"])
        g.view()

graph = 0

def opf(graph, naim, clustVar, engineVar, weightVar):
    def inopf():
        global graph, naim
        engine = engineVar.get()
        clust = clustVar.get()
        weight = weightVar.get()
        if clust == 1:
            cluster_umu = "yes"
        else:
            cluster_umu = "no"
        name = filedialog.askopenfilename(filetypes=[("Excel files", ".xls")])
        if name == "":
            return
        try:
            graph = Metamodule(name, cluster_umu = cluster_umu, engine = engine, weight = weight)
        except ImportError:
            messagebox.showerror("Could not make graph","Failed to make graph. Please check the file you chose.")
            return
        option(name)
        naim.destroy()
    return inopf

def intwo(graph, class1, class2, engine):
    def inintwo():
        print(engine.get())
        x = graph.classrelation(class1.get(), class2.get(), engine.get())
        if type(x) == str:    
            messagebox.showinfo("Class not found", "Please check class names and try again.")
        else:
            x.view()
    return inintwo
    
def two_classes():
    Twoclasses  = Tk()
    Twoclasses.title("Focus on two classes")
    engine = StringVar()
    Label(Twoclasses,text = "Enter class 1:").grid(row = 0, column = 0)
    class1 = Entry(Twoclasses)
    class1.grid(row = 0, column = 1)
    Label(Twoclasses,text = "Enter class 2:").grid(row = 1, column = 0)
    class2 = Entry(Twoclasses)
    class2.grid(row = 1, column = 1)
    engines = ["dot", "neato", "fdp", "sfdp", "twopi", "circo"]
    Label(Twoclasses, text = "Choose an engine:").grid(row = 2)
    for i, enginename in enumerate(engines):
        Radiobutton(Twoclasses, text = enginename, value = enginename, variable = engine).grid(row = 3 + i, column = 1)
    engine.set("dot")
    Button(Twoclasses, text = "View", command = intwo(graph, class1, class2, engine)).grid()

def save_source(root):
    def insave():
        save_name = filedialog.asksaveasfilename(parent = root)
        print("Path chosen:" + save_name)
        if save_name.endswith(".svg"):
            save_name = save_name[:-4]
        if save_name != "":
            graph.graph.render(save_name, format = "svg", cleanup = True)
            with open((save_name + "-source.txt"), "w") as src:
                src.write(graph.graph.source)
            messagebox.showinfo("Saved successfully", "Source file and svg file saved to " + save_name[:save_name.rfind("/")], parent = root )
        else:
            return
    return insave

def infocus(instance, engine):
    def innerfocus():
        graph.focus_input(instance.get(), engine.get())
    return innerfocus

def focusing():
    Focusdialog = Tk()
    Label(Focusdialog, text = "Enter the instance you want to see:").pack()
    instance = Entry(Focusdialog)
    instance.pack()
    engines = ["dot", "neato", "fdp", "sfdp", "twopi", "circo"]
    engine = StringVar()
    Label(Focusdialog, text = "Choose an engine:").pack()
    for i, enginename in enumerate(engines):
        Radiobutton(Focusdialog, text = enginename, value = enginename, variable = engine).pack()
    engine.set("dot")
    Button(Focusdialog, text = "Search", command = infocus(instance, engine)).pack()
    
    
def option(name):
    options = Tk()
    options.title("Choose an option")
    Label(options, text = f"Base file at {name} was analyzed successfully.\n").pack()
    Label(options, text = "What do you want to do?\n").pack()
    Button(options, text = "View Graph", command = graph.view).pack(fill = X, padx = 100)
    Button(options, text = "View Class Diagram", command = graph.class_diagram().view).pack(fill = X, padx = 100)
    Button(options, text = "Focus on two classes", command = two_classes).pack(fill = X, padx = 100)
    Button(options, text = "Save graph and source", command = save_source(options)).pack(fill = X, padx = 100)
    Button(options, text = "Focus on a single instance", command = focusing).pack(fill = X, padx = 100)

#overlap="false", cluster_umu="yes", cluster_samerank = "no"    
    
naim = Tk()
naim.title("Cobol analysis")
Label(naim, text = "Please choose the excel file with analysis results.").pack()
clust = IntVar()
engine = StringVar()
weight = IntVar()
Checkbutton(naim, text = "Use clusters", variable = clust).pack()
Checkbutton(naim, text = "Use weight for node sizes", variable = weight).pack()

engines = ["dot", "neato", "fdp", "sfdp", "twopi", "circo"]
Label(naim, text = "Choose an engine:").pack()
for i, enginename in enumerate(engines):
    Radiobutton(naim, text = enginename, value = enginename, variable = engine).pack()
engine.set("dot")
    
Button(naim, text = "Choose file...", command = opf(graph, naim, clust, engine, weight)).pack()
naim.mainloop()


