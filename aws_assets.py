import requests
from bs4 import BeautifulSoup as bs
import urllib.parse
import time
from datetime import datetime, timedelta
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import boto3
from base64 import b64decode

deepsecurity_key = '-'
deepsecurity_url = 'https://-/api/computers'
access_key = '-'
secret_key = '-'
scope = ['https://www.googleapis.com/auth/drive']
gcredentials = ServiceAccountCredentials.from_json_keyfile_name('-.json', scope)
gc = gspread.authorize(gcredentials)
# sh = gc.create('AWS List')
sh = gc.open('AWS List')

d = datetime.utcnow() - timedelta(hours=-9)
d = d.strftime('%Y%m%d')

# print(sh.id)
# worksheet = sh.worksheet('20190101-Server') # 20190101-Server
# worksheet2 = sh.worksheet('20190101-DB') # 20190101-DB

worksheet2 = sh.duplicate_sheet(source_sheet_id=0, insert_sheet_index=0, new_sheet_name=str(d)+'-DB')
worksheet2.clear()
RDS_cell_list = worksheet2.range('A1:Y1')
RDS_cell_list[0].value = 'Type'
RDS_cell_list[1].value = 'VPC Name'
RDS_cell_list[2].value = 'VPC ID'
RDS_cell_list[3].value = 'Team'
RDS_cell_list[4].value = 'State'
RDS_cell_list[5].value = 'CIDR Block'
RDS_cell_list[6].value = 'EndPoint'
RDS_cell_list[7].value = 'DB Identifier'
RDS_cell_list[8].value = 'Port'
RDS_cell_list[9].value = 'AvailabilityZone'
RDS_cell_list[10].value = 'Public Access'
RDS_cell_list[11].value = 'Engine'
RDS_cell_list[12].value = 'Engine Version'
RDS_cell_list[13].value = 'Storage'
RDS_cell_list[14].value = 'Multi AZ'
RDS_cell_list[15].value = 'vCPU'
RDS_cell_list[16].value = 'RAM'
RDS_cell_list[17].value = 'Instance Class'
RDS_cell_list[18].value = 'AutoMinorVersionUpgrade'
RDS_cell_list[19].value = 'Snapshot'
RDS_cell_list[20].value = 'DB Detail'
RDS_cell_list[21].value = 'Launch Time'
RDS_cell_list[22].value = 'Use'
RDS_cell_list[23].value = '-'
RDS_cell_list[24].value = '-'
worksheet2.update_cells(RDS_cell_list)

worksheet = sh.duplicate_sheet(source_sheet_id=0, insert_sheet_index=0, new_sheet_name=str(d)+'-Server')
worksheet.clear()
EC2_cell_list = worksheet.range('A1:Y1')
EC2_cell_list[0].value = 'Type'
EC2_cell_list[1].value = 'VPC Name'
EC2_cell_list[2].value = 'VPC ID'
EC2_cell_list[3].value = 'Team'
EC2_cell_list[4].value = 'VPC State'
EC2_cell_list[5].value = 'CIDR Block'
EC2_cell_list[6].value = 'Instance Name'
EC2_cell_list[7].value = 'Instance ID'
EC2_cell_list[8].value = 'Instance Type'
EC2_cell_list[9].value = 'AvailabilityZone'
EC2_cell_list[10].value = 'Instance State'
EC2_cell_list[11].value = 'Public DNS'
EC2_cell_list[12].value = 'Public IP'
EC2_cell_list[13].value = 'Key Name'
EC2_cell_list[14].value = 'Launch Time'
EC2_cell_list[15].value = 'Security Group'
EC2_cell_list[16].value = 'VPC ID'
EC2_cell_list[17].value = 'Private DNS'
EC2_cell_list[18].value = 'Private IP'
EC2_cell_list[19].value = 'EB Configuration'
EC2_cell_list[20].value = 'Image ID'
EC2_cell_list[21].value = 'OS Type'
EC2_cell_list[22].value = 'OS Detail'
EC2_cell_list[23].value = 'Use'
EC2_cell_list[24].value = 'IPS Version'
worksheet.update_cells(EC2_cell_list)

# sh.share('yeonghoon2@gmail.com', perm_type='user', role='writer')
def EC2_List(region):
    ec2_client = boto3.client(
        'ec2',
        # Hard coded strings as credentials, not recommended.
        aws_access_key_id = access_key,
        aws_secret_access_key = secret_key,
        region_name = region
    )
    ec2_resource = boto3.resource(
        'ec2',
        # Hard coded strings as credentials, not recommended.
        aws_access_key_id = access_key,
        aws_secret_access_key = secret_key,
        region_name = region
    )
    
    write_request = 1 # 100초이내에 100개의 write request만 가능해서,,
    response = ec2_client.describe_instances()
    for i in response['Reservations']:
        for j in i['Instances']:
            if write_request%35 == 0:
                # print('stop for 90sec')
                time.sleep(90)

            values_list = worksheet.col_values(8) # Instance ID column
            cell_list_index = 'A'+str(len(values_list)+1)+':Y'+str(len(values_list)+1)
            # print(cell_list_index)

            cell_list = worksheet.range(cell_list_index) # A ~ V column까지 작성
                        
            # worksheet.update_cells(cell_list)
            
            cell_list[0].value = 'Server'
            tmp_vpc_name = ''
            tmp_vpc_team = ''
            tmp_ec2_name = ''            
            # print(j['InstanceId'])
            tmp_ec2_id = j['InstanceId']
            instance = ec2_resource.Instance(tmp_ec2_id)
            tmp_ec2 = instance.tags
            for k in tmp_ec2:
                if k.get('Key') == 'Name':
                    tmp_ec2_name = k.get('Value')
            if instance.vpc_id != None:        
                vpc = ec2_resource.Vpc(instance.vpc_id)
                tmp_vpc = vpc.tags
                for k in tmp_vpc:
                    if k.get('Key') == 'Name':
                        tmp_vpc_name = k.get('Value')
                    if k.get('Key') == 'Team':
                        tmp_vpc_team = k.get('Value')

                cell_list[1].value = tmp_vpc_name # VPC Name
                cell_list[2].value = vpc.vpc_id # VPC ID
                cell_list[3].value = tmp_vpc_team # Team
                cell_list[4].value = vpc.state # VPC State
                cell_list[5].value = vpc.cidr_block # CIDR Block
            
            # worksheet.update_cell(len(values_list)+1, 1, 'Server') # Type
            cell_list[6].value = tmp_ec2_name # EC2 Name
            cell_list[7].value = tmp_ec2_id # Instance ID
            cell_list[8].value = instance.instance_type # Instance Type
            cell_list[9].value = instance.placement['AvailabilityZone'] # AvailabilityZone
            cell_list[10].value = instance.state['Name'] # Instance State
            cell_list[11].value = instance.public_dns_name # Public DNS(IPv4)
            cell_list[12].value = instance.public_ip_address # IPv4 Public IP
            cell_list[13].value = instance.key_name # Key Name
            cell_list[14].value = str(instance.launch_time) # Launch Time
            tmp_ec2_sg = instance.security_groups
            tmp_ec2_sg_str = ''
            for a, b in enumerate(tmp_ec2_sg):
                tmp_ec2_sg_str = tmp_ec2_sg_str + b['GroupName']
                if a < len(tmp_ec2_sg)-1:
                    tmp_ec2_sg_str = tmp_ec2_sg_str + '\n'
            cell_list[15].value = tmp_ec2_sg_str # SecurityGroup
            cell_list[16].value = instance.vpc_id # VPC ID
            cell_list[17].value = instance.private_dns_name # Private DNS
            cell_list[18].value = instance.private_ip_address # Private IP
            image = ec2_resource.Image(instance.image_id)
            try:
                cell_list[19].value = image.image_location # EB Configuration
            except:
                cell_list[19].value = '-' # EB Configuration
            cell_list[20].value = instance.image_id # AMI ID
            if instance.platform == None:
                os_type = '{}\n {}'.format('linux', instance.architecture)
            else:
                os_type = '{} {}'.format(instance.platform, instance.architecture)
            cell_list[21].value = os_type # OS type
            
            worksheet.update_cells(cell_list)
            write_request += 1
            
def RDS_List(region): 
    ec2_resource = boto3.resource(
        'ec2',
        # Hard coded strings as credentials, not recommended.
        aws_access_key_id = access_key,
        aws_secret_access_key = secret_key,
        region_name = region
    )    
    rds_client = boto3.client(
        'rds',
        aws_access_key_id = access_key,
        aws_secret_access_key = secret_key,
        region_name = region
    )
    
    write_request = 1
    response = rds_client.describe_db_instances()
    for i in response['DBInstances']:
        if write_request%35 == 0:
                time.sleep(90) # 3~4줄 적고(약 90개 request, 40초이상 소요), 1분 sleep -> 100초 뒤에 다시 동작)

        values_list = worksheet2.col_values(7) # Instance ID column
        cell_list_index = 'A'+str(len(values_list)+1)+':Y'+str(len(values_list)+1)
        # print(cell_list_index)

        cell_list = worksheet2.range(cell_list_index) # A ~ V column까지 작성

        # worksheet2.update_cells(cell_list)

        cell_list[0].value = 'DataBase'
        vpc_id = i['DBSubnetGroup']['VpcId']
        vpc = ec2_resource.Vpc(vpc_id)
        tmp_vpc_name = ''
        tmp_vpc_team = ''
        
        tmp_vpc = vpc.tags
        for k in tmp_vpc:
            if k.get('Key') == 'Name':
                tmp_vpc_name = k.get('Value')
            if k.get('Key') == 'Team':
                tmp_vpc_team = k.get('Value')
        cell_list[1].value = tmp_vpc_name
        cell_list[2].value = vpc_id
        cell_list[3].value = tmp_vpc_team
        cell_list[4].value = i['DBInstanceStatus']
        cell_list[5].value = vpc.cidr_block
        cell_list[6].value = i['Endpoint']['Address']        
        cell_list[7].value = i['DBInstanceIdentifier']
        cell_list[8].value = i['Endpoint']['Port']
        cell_list[9].value = i['AvailabilityZone']
        cell_list[10].value = i['PubliclyAccessible']
        cell_list[11].value = i['Engine']
        cell_list[12].value = i['EngineVersion']
        cell_list[13].value = str(i['AllocatedStorage']) + 'GB'
        cell_list[14].value = i['MultiAZ']
        cell_list[15].value = '-' # vCPU
        cell_list[16].value = '-' # RAM
        cell_list[17].value = i['DBInstanceClass']
        cell_list[18].value = i['AutoMinorVersionUpgrade']
        cell_list[19].value = i['CopyTagsToSnapshot']
        cell_list[20].value = '-' # DB 상세버전
        cell_list[21].value = str(i['InstanceCreateTime'])
        
        worksheet2.update_cells(cell_list)
        write_request += 1
        
def IPS_List():
    values_list = worksheet.col_values(8)
    cell_list_index = 'Y'+str(len(values_list)+1)+':Z'+str(len(values_list)+1)
    private_ip_list = worksheet.col_values(19)

    header = {
        'api-secret-key':deepsecurity_key,
        'api-version':'v1'
    }
    r = requests.get(deepsecurity_url, headers=header, verify=False)
    result = r.json()  
    
    for c, i in enumerate(private_ip_list):
        for j in result['computers']:
            privateIPAddress = j['ec2VirtualMachineSummary']['privateIPAddress']
            agentVersion = j['agentVersion']
            if agentVersion == '0.0.0.0':
                continue
            if privateIPAddress == i:
                # print(i, privateIPAddress)
                worksheet.update_cell(c+1, 25, agentVersion)
                time.sleep(10)

# 'us-east-1' # N.Virginia
# 'ap-northeast-2' # Seoul 
EC2_List('ap-northeast-2')
EC2_List('us-east-1')
RDS_List('ap-northeast-2')
RDS_List('us-east-1')
IPS_List()                