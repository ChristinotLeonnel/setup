from utils import * 
import matplotlib.pyplot as plt

def XLABEL(span:bool=True ,neoud:bool=True, grid:bool=True):
    plt.clf()
    ax = plt.gca()
    ax.set_xticks(neouds)
    ax.set_xticklabels(distances)
    plt.xticks(rotation=45)
    if span:
        ax.plot(x_normal, [0]*len(x_normal), 'k--')  
    if neoud:
        ax.scatter(neouds, [0]*len(neouds))
    if grid:
        ax.grid(True)  

def SOLO(lines:str ="Bending Moment.json" , span:int=0 ,section:int=0 ,legend:bool=True):
    XLABEL()
    if lines in ["Deflection.json","Bending Moment.json","Rotation.json"]:
        plt.plot(x_normal ,open_json(lines)[span][section], label=f"Travee : {span+1}\nSection : {section}") 
    elif lines == "Shear Force.json":
        plt.plot(x_forces[span][section] ,open_json(lines)[span][section], label=f"Travee : {span+1}\nSection : {section}")
    elif lines == "Support Moment.json":
        plt.plot(x_normal, open_json(lines)[span], label=f"M_{span}") 
    elif lines == "support_reactions":
        plt.plot(x_forces[span][0], open_json(lines)[span], label=f"R_{span}") 

    if legend:
        plt.title(lines.split(".")[0].upper(), text_format) 
        plt.xlabel("Distance entre les appuis".upper(), text_format) 
        plt.ylabel("values of ".upper() + list(lines.split("."))[0].upper(), text_format)   
        plt.legend()

def SPAN_FORCES(span:int=None, section:int=None ,legend:bool=True):
    XLABEL() 
    FORCE = open_json("Shear Force.json")
    lines = "Shear Force.json"

    if span==None and section==None:
        for i,j in zip(x_forces, FORCE):
            for a,b in zip(i,j):
                plt.plot(a,b) 
        if legend:
            plt.xlabel("Distance entre les appuis".upper(), text_format) 
            plt.ylabel("values of ".upper() + list(lines.split("."))[0].upper(), text_format)  
            plt.title("ALL SPAN_SECTION")  

    elif span!=None and section==None:
        for i,j in zip(x_forces[span], FORCE[span]):
            plt.plot(i,j)
        plt.title(f"ALL SECTION FOR SPAN_{span+1}") 
        plt.xlabel("Distance entre les appuis".upper(), text_format) 
        plt.ylabel("values of ".upper() + list(lines.split("."))[0].upper(), text_format)  

    elif span==None and section!=None:
        for i in range(len(neouds)-1):
            plt.plot(x_forces[i][section], FORCE[i][section],label = f"Travee {i} :: section {section}")
        if legend:
            plt.xlabel("Distance entre les appuis".upper(), text_format) 
            plt.ylabel("values of ".upper() + list(lines.split("."))[0].upper(), text_format)  
            plt.legend() 
        plt.title(f"ALL TRAVEE FOR SECTION {section}") 

    else:
        SOLO("Shear Force.json", span, section, legend) 

def SPAN_Three_boys(lines:str="Bending Moment.json", span:int=None, section:int=None ,legend:bool=True):
    XLABEL() 
    FORCE = open_json(lines)

    if span==None and section==None:
        for i in FORCE:
            for j in i:
                plt.plot(x_normal,j) 
        if legend:
            plt.xlabel("Distance entre les appuis".upper(), text_format) 
            plt.ylabel("values of ".upper() + list(lines.split("."))[0].upper(), text_format)  
            plt.title("ALL SPAN_SECTION")  

    elif span!=None and section==None:
        for i in FORCE[span]:
            plt.plot(x_normal , i)
        plt.title(f"ALL SECTION FOR SPAN_{span+1}") 
        plt.xlabel("Distance entre les appuis".upper(), text_format) 
        plt.ylabel("values of ".upper() + list(lines.split("."))[0].upper(), text_format)  

    elif span==None and section!=None:
        for i in range(len(neouds)-1):
            plt.plot(x_normal, FORCE[i][section],label = f"Travee {i} :: section {section}")
        if legend:
            plt.xlabel("Distance entre les appuis".upper(), text_format) 
            plt.ylabel("values of ".upper() + list(lines.split("."))[0].upper(), text_format)  
            plt.legend() 
        plt.title(f"ALL TRAVEE FOR SECTION {section}") 

    else:
        SOLO(lines, span, section, legend) 

def SPAN_Support(span:int=None, legend:bool=True, first_end:bool=True): 
    XLABEL() 
    lines = "Support Moment.json"
    FORCE = open_json(lines) 

    if span == None:
        if not first_end:
            for a,i in enumerate(FORCE):
                plt.plot(x_normal, i, label=f"M_{a}") 
        else:
            for i in range(1,len(neouds)-1):
                plt.plot(x_normal, FORCE[i], label=f"M_{i}") 

        if legend:
            plt.legend() 
    else:
        SOLO(lines, span, legend=legend)

def MAX(lines:str="Bending Moment.json", area:bool=False):
    max_M_point = open_json(f"{lines.split(".")[0]} Max Positions.json", "Analysis")
    if lines=="Support Moment.json":
        if area:
            SOLO(lines, span=max_support_areas) 
        else:
            SOLO(lines, span=max_M_point["span"]) 
    else:
        if area:
            SOLO(lines, span=max_M_areas["travee"], section=max_M_areas["section"])
        else:
            SOLO(lines, span=max_M_point["span"], section=max_M_point["section"])

def TRACE(lines:str ="Bending Moment.json" , span:int=None ,section:int=None ,legend:bool=True, 
          max:bool=False, area:bool=True, first_end:bool=True):
    if max:
        MAX(lines, area)
    elif lines == "Shear Force.json":
        SPAN_FORCES(span, section, legend)
    elif lines == "Support Moment.json":
        SPAN_Support(span, legend, first_end)
    elif lines in ["Deflection.json","Bending Moment.json","Rotation.json"]:
        SPAN_Three_boys(lines, span, section, legend) 
    elif lines == "support_reactions":
        if span==None:
            span = 0
        SOLO(lines, span=span,legend=legend)

TRACE("Bending Moment.json", max=True, area = False)
plt.show() 
