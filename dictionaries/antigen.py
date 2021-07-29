from docxtpl import DocxTemplate
import cx_Oracle
import os
import json
from kelas.qr import QrCode

#SETUP TEST(S) HERE
params = ['SARSH']
#END SETUP TEST(S)

#SETUP TEMPLATE HERE
template = r'dictionaries\templates\antigen.docx'
#END SETUP TEMPLATE

#convert test(s) array to string
param_in_string = ','.join("'{}'".format(param) for param in params)


def getdata(filename, labno):
    #get database config
    setting_file = os.path.abspath(os.path.join(os.path.dirname(__file__),"..","setting.json"))
    with open(setting_file,'r') as f :
        config = json.load(f)
    user_db = config["database"]["user"]
    pass_db = config["database"]["pass"]
    host_db = config["database"]["host"]

    try:
        conn = cx_Oracle.connect(user_db,pass_db,host_db)
        cursor = conn.cursor()
        sql = """
            select 
            to_char(oh_trx_dt,'dd-MON-yyyy'), oh_tno, oh_pid, oh_apid, oh_last_name, 
            (case when oh_pataddr1 is null then '-' else oh_pataddr1 end), 
            (case when oh_pataddr2 is null then '-' else oh_pataddr2 end), 
            (case when oh_pataddr3 is null then '-' else oh_pataddr3 end), 
            (case when oh_pataddr4 is null then '-' else oh_pataddr4 end), 
            to_char(oh_bod,'dd-MON-yyyy'), to_char(os_spl_rcvdt,'dd-MON-yyyy hh24:mi'),
        """
        for param in params:
            sql = sql + "max(decode(od_testcode,'{}',(case tv_desc when null then od_tr_val else tv_desc end))){},".format(param,param)
        sql = sql + f"""
            (case oh_sex when '1' then 'Laki-laki' when '2' then 'Perempuan' end)oh_sex
            from hord_hdr
            join hord_dtl on oh_tno = od_tno
            join hord_spl on os_tno = oh_tno and od_spl_type = os_spl_type
            left join textvalue on tv_code = od_tr_val
            where oh_tno = '{labno}'
            and od_testcode in({param_in_string})
            group by oh_trx_dt, oh_tno, oh_pid, oh_apid, oh_last_name, oh_pataddr1, oh_pataddr2, oh_pataddr3, oh_pataddr4,
            oh_bod, oh_sex, os_spl_rcvdt
        """

        print(sql)
        
        cursor.execute(sql)
        record = cursor.fetchone()

        if record is not None:
            trx_dt = record[0]
            lno = record[1]
            pid = record[2]
            apid= record[3]
            name= record[4]
            addr1= record[5]
            addr2= record[6]
            addr3= record[7]
            addr4= record[8]
            dob = record[9]
            specimen_on = record[10]

            #GET PARAM VALUE HERE
            swbpcr = record[11]

            #DO NOT MOVE SEX FIELD TO UPPER SECTION
            sex = record[12]
        
        else:
            trx_dt = lno = pid = apid = name = addr1 = addr2 = addr3 = addr4 = dob = specimen_on = swbpcr = sex = ''


        #SETUP VARIABLE VALUE HERE
        replacements = {
            'trx_dt' : trx_dt,
            'apid': apid,
            'nama': name,
            'address1' : addr1,
            'address3' : addr2,
            'address4' : addr4,
            'dob': dob,
            'value': swbpcr,
            'sex': sex
            }
        #END SETUP VARIABLE VALUE
        
        #GENERATE QR CODE
        context = f"""
            Hasil Laboratorium : RSUD Leuwiliang
            No Lab      : {lno}
            Nama        : {name}
            Tgl Lahir   : {dob}
            Tgl Spesimen: {specimen_on}
            Nama Test   : Antigen SARS-CoV-2
            Hasil       : {swbpcr}
            Status      : VERIFIED
        """

        #get word data and replace specific word based on replacement variable
        doc = DocxTemplate(template)

        qr = QrCode(context)
        qr.save(lno)

        doc.replace_pic('Picture 3',os.path.join('temp',lno+'.jpg'))

        doc.render(replacements)

        doc.save(filename)

        qr.delete(lno)
        
    except cx_Oracle.Error as e:
        print(e)

if __name__ == "__main__":
    getdata("D:\\test.docx","21000001")