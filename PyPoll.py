import csv

csv_path = "C:\\Users\\nupur_v6\\01-Lesson-Plans\\03-Python\\3\\Python-Challenge HW\\Resources\\election_data.csv"
output_file = "C:\\Users\\nupur_v6\\01-Lesson-Plans\\03-Python\\3\\Python-Challenge HW\\Resources\\PyPoll output file.txt"

total_count = 0; 
khan_count = 0; 
correy_count = 0; 
li_count = 0; 
otooley_count = 0; 
winner_votecount = 0

def percentage (part, whole):
    return 100 * float(part)/float(whole)

with open(csv_path, newline='') as csvfile:
     csvreader = csv.reader(csvfile, delimiter=',')

     for i in csvreader:
         voter_id = i[0]
         county = i[1]
         candidate = i[2]
    
         total_count = total_count + 1

         if candidate =="Khan":
            khan_count = khan_count + 1
         if candidate =="Correy":
            correy_count = correy_count + 1
         if candidate =="Li":
            li_count = li_count + 1
         if candidate =="O'Tooley":
            otooley_count = otooley_count + 1        

     candidatevote = {"Khan": khan_count,"Correy": correy_count,"Li" :li_count, "O'Tooley": otooley_count}
     for candidate, value in candidatevote.items():
         if value > winner_votecount:
            winner_votecount = value
            winner = candidate
output = (     
    f'Election Results'+'\n'
    f'-------------------------------'+'\n'
    f'Total Votes: {total_count}'+'\n'
    f'-------------------------------'+'\n'
    f'Khan: {percentage(khan_count,total_count):.3f}%  ({khan_count})'
    f'Correy: {percentage(correy_count,total_count):.3f}%  ({correy_count})'
    f'Li: {percentage(li_count,total_count):.3f}%  ({li_count})'
    f'O\'Tooley: {percentage(otooley_count,total_count):.3f}%  ({otooley_count})'
    f'--------------------------------'+'\n'
    f'Winner: {winner} '+'\n'
    f'--------------------------------'+'\n')
print(output)
with open(output_file, "w") as txt_file:
    txt_file.write(output)