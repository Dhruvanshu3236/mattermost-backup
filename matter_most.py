import requests
import json
import os

from datetime import datetime
import xlsxwriter


url = "https://mm.upforce.tech"
auth_token = ""
login_url = url+"/api/v4/users/login"

payload = { "login_id": "dhruvanshu@upforce.tech",
            "password": "Dhruvanshu@1775"}
headers = {"content-type": "application/json"}
s = requests.Session()
r = s.post(login_url, data=json.dumps(payload), headers=headers)

id = r.json()
user_id = id['id']

auth_token = r.headers.get("Token")

hed = {'Authorization': 'Bearer ' + auth_token}

channel_name = "qa-team"




if not os.path.exists(channel_name):
    os.makedirs(channel_name)



def get_channel_id():

    channel_url_main = url + '/api/v4/channels' 
    response = requests.get(channel_url_main, headers=hed)
    info1 = response.json()

    channel_id = ""

    for i in info1:
        if i['name'] == channel_name:
            channel_id = i['id']



    team_url_1 = url+'/api/v4/channels/'+channel_id+'/posts?per_page=100000000&include_deleted=true'
    response = requests.get(team_url_1, headers=hed)
    info = response.json()
    get_all_post_info(info)



def get_all_post_info(info):

    main_all_time_chat_data = {}

    for i, v in info['posts'].items():


       
        user_url = url + f'/api/v4/users/{v["user_id"]}'
        response = requests.get(user_url, headers=hed)
        user_info = response.json()


        dt = datetime.fromtimestamp(int(v['create_at']) / 1000)
        formatted_time = dt.strftime('%d-%m-%Y %H:%M:%S.%f')[:-3]

        time_op = formatted_time[0:10]

        if not os.path.exists(channel_name+'/'+formatted_time[0:10]):
            os.makedirs((channel_name+'/'+formatted_time[0:10]))
        

        if time_op not in main_all_time_chat_data:
            main_all_time_chat_data[time_op] = []

        mm_data = {"id":i,"Username":user_info['username'],"Time":formatted_time,"Message":v['message']}


        main_all_time_chat_data[time_op] += [mm_data] 


    save_all_message_file(main_all_time_chat_data)


def save_all_message_file(data):
    for i,v in data.items():
        create_csv_for_message(i,v)
        
        for k in v:
            post_url = url+'/api/v4/posts/'+k['id']+'/files/info'
            response = requests.get(post_url, headers=hed)
            info = response.json()


            try:
                if info != [] or 'id':
                    file_val = 1 
                    for x in info:
                        id = x['id']

                        file_url = url+'/api/v4/files/'+id
                        response = requests.get(file_url, headers=hed, )

                        path_full = channel_name+'/'+str(i)

                        open(os.path.join(path_full, str(str(file_val)+str(info[0]['name']))), 'wb').write(response.content)

                        file_val += 1
            except:
                pass

def create_csv_for_message(date, content):

        
    file_name = f"{channel_name}/{date}/chat.xlsx"

    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold': True})

    header = workbook.add_format({'size':18,'bold':True})

    worksheet.set_column(0, 4, 20)
    worksheet.set_row(1, 25)


    worksheet.write('B2', 'Message History',header)


    worksheet.write('A4', 'Username', bold)
    worksheet.write('B4', 'Date', bold)
    worksheet.write('C4', 'Message', bold)



    row = 4
    column = 0
    
    

    for item in content :
        worksheet.write(row, column, item['Username'])
        worksheet.write(row, column + 1, item['Time'])
        worksheet.write(row, column + 2, item['Message'])
    

    
        row += 1
        
    workbook.close()   


        


channel = get_channel_id()
