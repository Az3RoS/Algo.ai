import itertools as it
import xlrd 
import copy

def average_rank(w_list):    
    rank_sum = 0
    dcount = 0
    rank_list = [0 for x in range(len(w_list))] 
    list_x = sorted(range(len(w_list)),key=w_list.__getitem__)
    list_y = [w_list[pos] for pos in list_x]    
    
    for i in range(len(w_list)):
        rank_sum = rank_sum+i
        dcount = dcount+1
        if i==len(w_list)-1 or list_y[i] != list_y[i+1]:
            for j in range(i-dcount+1,i+1):
                rank_avg = rank_sum / float(dcount) + 1
                rank_list[list_x[j]] = rank_avg
            rank_sum = 0
            dcount = 0 
            
    return rank_list

def wilcoxon_signed_rank_test(data_dict_w):
    
    '''
    Wilcoxon Signed Rank Test statistic - Wikipedia and Other sources 
    '''
    data_cols = list(data_dict_w.keys())
    required_cols = data_cols[2:]
    dict_pair_wiki = {}
    dict_pair_theory = {}
    
    for i in range(len(required_cols)):
        rank_list = []
        wtest_dict = {}
        wtest_dict['key'] = data_dict_w[data_cols[1]]
        wtest_dict['nonkey'] = data_dict_w[data_cols[i+2]]
        
        wtest_dict['diff'] = [((wtest_dict['key'][i]-wtest_dict['nonkey'][i])) for i in range(len(wtest_dict['key'])) if ((wtest_dict['key'][i]-wtest_dict['nonkey'][i]) !=0)]
        wtest_dict['abs_diff'] = [(abs(wtest_dict['diff'][i])) for i in range(len(wtest_dict['diff']))]
        wtest_dict['sgn'] = [1 if (wtest_dict['diff'][i])>0 else -1 for i in range(len(wtest_dict['diff']))]
        rank_list = average_rank(wtest_dict['abs_diff'])      
        wtest_dict['rank'] = rank_list
        
        #ranked from lowest to highest, with tied ranks included where appropriate
        wtest_dict['sgn_x_rank'] =  [(wtest_dict['sgn'][i])*(wtest_dict['rank'][i]) for i in range(len(wtest_dict['diff']))]         
        dict_pair_wiki.update({(data_cols[1],data_cols[i+2]):sum(wtest_dict['sgn_x_rank'])})
        
        wtest_dict['neg_rank'] = [(wtest_dict['sgn_x_rank'][i]) for i in range(len(wtest_dict['diff'])) if (wtest_dict['sgn_x_rank'][i]) <0 ]
        wtest_dict['pos_rank'] = [(wtest_dict['sgn_x_rank'][i]) for i in range(len(wtest_dict['diff'])) if (wtest_dict['sgn_x_rank'][i]) >0 ]  
        nrank_sum = abs(sum(wtest_dict['neg_rank']))
        prank_sum = sum(wtest_dict['pos_rank'])
       
        # Wikipedia - W stat is equal to sum of all singed ranks 
        dict_pair_wiki.update({(data_cols[1],data_cols[i+2]):sum(wtest_dict['sgn_x_rank'])})
        
        # Other sources - W stat is equal to min of sum of negtive ranks and sum of positive ranks
        dict_pair_theory.update({(data_cols[1],data_cols[i+2]):min(nrank_sum,prank_sum)})
       
    
    return (dict_pair_wiki,dict_pair_theory) 

def permu_n_weights(data_dict_p):
    
    '''
    Calculate all permutations of the non-key columns 
    '''    
    comb_dict={}
    comb_dict=copy.deepcopy((data_dict_p))
    all_columns = list(comb_dict.keys())
    nonkey_cols = all_columns[2:]
    prod_combinations = []
    
    for i in range(len(nonkey_cols)-1):
        prod_combinations.extend(list(it.combinations(nonkey_cols, i+2)))  
    
    for element in prod_combinations:
        sum=[0 for x in range(len(data_dict_p[all_columns[1]]))] 
        col_name = "W_"
        for sub in element:
            S = sub
            sum=[x+y for x,y in zip(sum, comb_dict[S])]
            col_name = col_name + str(sub) + "_"
        col_name = col_name[:-1]
        comb_dict[col_name] = [ val/len(element)  for val in sum] 
    
    for ele in nonkey_cols:
        del comb_dict[ele] 
        
    return (prod_combinations,comb_dict)   


def main():
     
    loc = (r'C:\Python Scripts\data\Algo Test.xlsx') 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheets()[0]
    sheet = wb.sheet_by_name("Data")
    sheet = wb.sheet_by_index(0)  
    data_dict = {}
    pos=0
    col_names = sheet.row_values(0)
    for element in sheet.row_values(0):
        data_dict.update({element:sheet.col_values(pos)[1:]})
        pos += 1
    
    
    dict_pair_wiki,dict_pair_theory = wilcoxon_signed_rank_test(data_dict)
    
    print("\n\nAnswer to Question 1...")
    print("\nKey and Non Key product pair with W test (Wikipedia Defination):\n")
    for index, (key, value) in enumerate(dict_pair_wiki.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    print("\nKey and Non Key product pair with W test (Other Defination):\n")
    for index, (key, value) in enumerate(dict_pair_theory.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    prod_comb,PerW_dict = permu_n_weights(data_dict)
    pw_pair_wiki,pw_pair_theory = wilcoxon_signed_rank_test(PerW_dict)
    
    print("\n\nAnswer to Question 2...")
    print("\nNumber of Product Combinations:\n")
    print(len(prod_comb))
     
    print("\nProduct Combinations:\n")
    print((prod_comb))
    
    print("\nKey and Non key combination with W test (Wikipedia Defination):\n")
    for index, (key, value) in enumerate(pw_pair_wiki.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    print("\nKey and Non key combination with W test (Other Defination):\n")
    for index, (key, value) in enumerate(pw_pair_theory.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    bestfit_nonkey = min(dict_pair_wiki, key=lambda y: abs(dict_pair_wiki[y]))
    print("\n\nAnswer to Question 3...")
    print("\nBest Fit non key product:\n")
    print(bestfit_nonkey)      
                              
    
    print("\n\nAnswer to Question 4...")
    print("\nBest Fit non key product combinations:\n")
    bestfit_pernw = min(pw_pair_wiki, key=lambda y: abs(pw_pair_wiki[y]))
    print(bestfit_pernw)     


if __name__ == "__main__":
    main()