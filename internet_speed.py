import speedtest , openpyxl , time 
from math import ceil


while True:

    test = speedtest.Speedtest()

    print("Loading server list...")
    test.get_servers()
    print("finding best server...")
    best = test.get_best_server()
    print(f"Found: {best['host']} located in {best['country']}")


    print("performing download test...")
    download_result = "{:.2f}".format(test.download() / 1024 / 1024)
    print("performing upload test...")
    upload_result ="{:.2f}".format(test.upload() / 1024 / 1024) 
    print("performing ping test...")
    ping_result ="{:.2f}".format(test.results.ping)

    print( "Download speed: "+ download_result +" Mbit/s")
    print( "Upload speed: "+ upload_result +" Mbit/s")
    print( "Ping: "+ ping_result +" ms")

    #data base side

    excelfile = openpyxl.load_workbook("speedtest_result.xlsx")
    sheet1 = excelfile.active

    now = time.localtime()
    for x in range(1,999999):
            if sheet1["A" + str(x)].value == None:
                sheet1["A" + str(x)].value = now[3]
                sheet1["B" + str(x)].value = download_result
                sheet1["C" + str(x)].value = upload_result
                sheet1["D" + str(x)].value = ping_result
                break

    #excelfile.save("speedtest_result.xlsx")
    
    
    time.sleep(60*30)