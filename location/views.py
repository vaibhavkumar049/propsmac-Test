from django.shortcuts import render
from django.http import HttpResponse
from django.contrib import messages
import pandas as pd
import googlemaps
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import xlwt
# Create your views here.


def index(request):
    if request.method=="GET":
        prompt = {
        'order': "Excel file should have 'Address1','Address3','Address4','Address5' field",
              }
        return render(request,'location/index.html',prompt)
    
    try:
        excel_file = request.FILES['file']    
        if not excel_file.name.endswith('.xlsx'):
            messages.error(request, 'THIS IS NOT A EXCEL FILE')
        data=pd.read_excel(excel_file,index_col=0)
        cols=['Address1','Address3','Address4','Address5']
        data['ADDRESS']=data[cols].apply(lambda row: ','.join(row.values),axis=1)

        locator = Nominatim(user_agent='test-app')
        geocode = RateLimiter(locator.geocode, min_delay_seconds=1)
        data['location'] = data['ADDRESS'].apply(geocode)
        data['point'] = data['location'].apply(lambda loc: tuple(loc.point) if loc else None)
        data[['latitude', 'longitude', 'altitude']] = pd.DataFrame(data['point'].tolist(), index=data.index)
        data.drop(['point','altitude','location'],axis=1,inplace=True)

        # gmap=googlemaps.Client(key='api-key')
        # geocode_res=gmap.geocode('Manora,Darbhanga,Bihar,India')
        # print("res",geocode_res[0]["geometry"]["location"])

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="address.xlsx"'

        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('address')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        columns=data.columns.values.tolist()
        for col_num in range(len(columns)):
            ws.write(row_num, col_num, columns[col_num], font_style)

        font_style = xlwt.XFStyle()
        for row in data.values.tolist():
            row_num += 1
            for col_num in range(len(row)):
                ws.write(row_num, col_num, row[col_num], font_style)
        
        wb.save(response)
        # render(request,'location/index.html',{'columns':data.columns.values.tolist(),'rows':data.to_dict('records')})
        return response
    except Exception as e:

        messages.error(request,"Unable to upload Excel file. "+repr(e))
        return render(request,'location/index.html')

        

