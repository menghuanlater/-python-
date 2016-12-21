#Max Clique

def DFS(VertexSets,layer):
    global MaxClique
    if(layer > 5):  #when the phrase is right,we must have a clique which maybe the best answer
        if(len(VertexSets) > MaxClique):
            MaxClique = len(VertexSets)
            print("-*_*-Current Max Clique is: [",end = "")
            #the print vertexes the sets conclude
            for i in range(0,len(VertexSets)-1):
                print(VertexSets[i],end = ", ")
            print(VertexSets[len(VertexSets)-1],"]")

        elif(len(VertexSets) == MaxClique):
            print("-*_*-Another Clique equals to Current Max Clique is: [",end = "")
            #the print vertexes the sets conclude
            for i in range(0,len(VertexSets)-1):
                print(VertexSets[i],end = ", ")
            print(VertexSets[len(VertexSets)-1],"]")
        return ; #back search

    if(layer == 1): #start the find
        VertexSets.append(layer-1) #add first vertex in 
        DFS(VertexSets,layer+1) #DFS 
        VertexSets.pop() #delete first vertex
        DFS(VertexSets,layer+1) #exeute same find_procedure

    else:
        flag = True
        for i in range(0,len(VertexSets)):
            if(RelationTrix[layer-1][VertexSets[i]]==0):
                flag = False #if new vertex can't append to the current clique,set flag to false
                break
        #judge whether layer-1 vertex can be appended into VertexSets
        if(flag):
             VertexSets.append(layer-1)
             DFS(VertexSets,layer+1)
             VertexSets.pop()
             DFS(VertexSets,layer+1)
        else:
            DFS(VertexSets,layer+1)

def main():
    global MaxClique
    MaxClique = 0
    #first init the relation matrix
    for i in range(0 , 5):
        for j in range(0 , 5):
            RelationTrix[i].append(0)
    #the deal the edge information to accomplish 
    for i in range(0 , 5):
        length = len(E[i])
        for j in range(0,length):
            RelationTrix[i][E[i][j]] = 1
            RelationTrix[E[i][j]][i] = 1
    #now we have create a graph
    #the next step is Find the Max Clique
    VertexSets = [] #Initial,this list is for vertexes in one clique
    DFS(VertexSets,1)#depth first search-->serach back to find best answer

if __name__ == "__main__":
    V = [0, 1, 2, 3, 4] #five vertex
    E = [[1, 3, 4], [2, 3, 4], [3, 4], [4], []]
    RelationTrix = [[], [], [], [], []]
    main()
