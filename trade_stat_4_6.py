#Trade Stat - scripting program for MOEX traders that use QUIK program to work
#It takes HTML report file (Source_fn) and  geterates output txt file that contains trading statistics^
# Buy and Sell average
# Dispersion


import math
import openpyxl
import operator
import datetime
import bs4

Source_fn="rep58076.htm"

def load_data_htm(data_arr_b,data_arr_s,names_b,names_s,filename):
    soup=bs4.BeautifulSoup(open(filename))
# the next line are for debugging
#    file=open("output.txt","wt")
    table = soup.find("table", attrs={"border":"1px"})
    i=0
    j=0
    k=0
    for row in table.find_all("tr")[1:]:
        dataset = [td.get_text() for td in row.find_all("td")]
        if len(dataset)<16:
            break
        k+=1
        if dataset[6]=="Продажа":
            data_arr_s.append([])
            names_s.append(dataset[3])
            data_arr_s[i].append(int(dataset[7].replace(" ", "")))
            data_arr_s[i].append(float(dataset[9].replace(" ", "")))
            i+=1
        else:
            data_arr_b.append([])
            names_b.append(dataset[3])
            data_arr_b[j].append(int(dataset[7].replace(" ", "")))
            data_arr_b[j].append(float(dataset[9].replace(" ", "")))
            j+=1
#next lines are for debugging
#        for ch in dataset:
#            file.write(ch+"|")
#        file.write("\n")
#    file.close()
    print("File loaded. Sells=",i,"; Buys=",j,"; Total records=",k)
  
#this procedure for debbuging purposes only
def save_data(nam_arr,data_arr,fn):
    file=open(fn,"wt")
    for i in range(len(nam_arr)):
        file.write(nam_arr[i]+"|"+str(data_arr[i][0])+"|"+str(data_arr[i][1])+"\n")
    file.close()

def build_index(tr_arr, ind_arr,names_array):
#names_array contains unique values
    for i in range(len(tr_arr)):
        try:
            k=names_array.index(tr_arr[i])
        except:
            names_array.append(tr_arr[i])
            ind_arr.append([])
            ind_arr[len(ind_arr)-1].append(i)
        else:
            ind_arr[k].append(i)

def find_avrg(data_arr, ind_arr, res_arr):
    for i in range(len(ind_arr)):
        amount=0
        price=0
        res_arr.append([])
        for j in range(len(ind_arr[i])):
            amount+=data_arr[ind_arr[i][j]][0]
            price+=data_arr[ind_arr[i][j]][1]
        res_arr[i].append(price/amount)
        res_arr[i].append(amount)

#in avr sq disperce calculation assumes that
#for single transaction for A shares and P price avrg would be A*(P-avrg_P)^2
    for i in range(len(ind_arr)):
        part_sum=0
        for j in range(len(ind_arr[i])):
            part_sum+=data_arr[ind_arr[i][j]][0]*(data_arr[ind_arr[i][j]][1]/data_arr[ind_arr[i][j]][0]-res_arr[i][0])*(data_arr[ind_arr[i][j]][1]/data_arr[ind_arr[i][j]][0]-res_arr[i][0])
        if (res_arr[i][1]==1):
            res_arr[i].append(0)
        else:
#            res_arr[i].append(math.sqrt(part_sum/(res_arr[i][1]-1))/res_arr[i][0])
            res_arr[i].append(math.sqrt(part_sum/(res_arr[i][1]-1)))


nam_array_b=[]
nam_array_s=[]
data_array_b=[]
data_array_s=[]
index_s=[]
index_b=[]
names_s=[]
names_b=[]
res_arr_s=[]
res_arr_b=[]

load_data_htm(data_array_b,data_array_s,nam_array_b,nam_array_s,Source_fn)

#next lines are for debugging
#save_data(nam_array_b,data_array_b,"out_buy.txt")
#save_data(nam_array_s,data_array_s,"out_sel.txt")

print("Building indexes... ",end="")
build_index(nam_array_s,index_s,names_s)
build_index(nam_array_b,index_b,names_b)

print("Completed. Total sells=",len(names_s)," buys=",len(names_b),"records.")

print("Calculating averages...")
find_avrg(data_array_s,index_s,res_arr_s)
find_avrg(data_array_b,index_b,res_arr_b)

#writing to the file
out_file=open(datetime.datetime.strftime(datetime.datetime.now(), 'stat%d%m%Y.txt'), "wt")
out_file.write(" #  | Stock | Avg Sell   |Dispers.| Amount |  Avg Buy   |Dispers.| Amount |\n")
for i in range(len(names_s)):
    try:
        out_file.write("%3d | %5s | %10.4f |%8.3f|%8d| %10.4f |%8.3f|%8d|\n" % (i,names_s[i],res_arr_s[i][0],res_arr_s[i][2],res_arr_s[i][1],res_arr_b[names_b.index(names_s[i])][0],res_arr_b[names_b.index(names_s[i])][2],res_arr_b[names_b.index(names_s[i])][1]))
    except:
        out_file.write("%3d | %5s | %10.4f |%8.3f|%8d| \n" % (i,names_s[i],res_arr_s[i][0],res_arr_s[i][2],res_arr_s[i][1]))
k=len(names_s)
for i in range(len(names_b)):
    try:
        names_s.index(names_b[i])
    except:
        out_file.write("%3d | %5s |            |        |        | %10.4f |%8.3f|%8d|\n" % (k,names_b[i],res_arr_b[i][0],res_arr_b[i][2],res_arr_b[i][1]))
        k+=1

out_file.close()
print("Calculation complete.")


