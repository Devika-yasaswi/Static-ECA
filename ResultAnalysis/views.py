from django.shortcuts import render
from django.http import FileResponse, JsonResponse
from utils.regular_SGPA import *
from utils.pdfToDataframe import *
from utils.Revaluation import *
from utils.branch_wise_analysis import *
from pandas import read_excel

# Create your views here.
def home(request):    
    return render(request, 'Home.html')
def resultAnalysis(request):
    return render(request, 'Result Analysis.html')
def seating(request):
    return render(request,'seatinghome.html') 
def process_regular_sgpa(request):
    if request.method == 'POST':
        # Get the uploaded file from the request
        regular_class_file = request.FILES.get('regular_class')
        selected_branch=request.POST.get('selected_branch')
        print(selected_branch)
        # Get the MIME type of the uploaded file
        file_mime_type = regular_class_file.content_type
        if file_mime_type == 'application/pdf':
            return_data=pdfToDataframe(regular_class_file)
            if isinstance(return_data, pd.DataFrame):
                value=Sgpa(return_data)
                if isinstance(value,str):
                    return JsonResponse({'message': value},safe=False)
                if selected_branch != None:
                    branchwise_analysis('Result.xlsx',selected_branch)
                response = FileResponse(open('Result.xlsx', 'rb'))
                response['Content-Disposition'] = 'attachment; filename="Result.xlsx"'
                return response
            elif isinstance(return_data,str):
                return JsonResponse({'message': return_data},safe=False)
        elif file_mime_type=='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or file_mime_type=='application/vnd.ms-excel':
            df=read_excel(regular_class_file)
            value=Sgpa(df)
            if isinstance(value,str):
                return JsonResponse({'message': value},safe=False)
            response = FileResponse(open('Result.xlsx', 'rb'))
            response['Content-Disposition'] = 'attachment; filename="Result.xlsx"'
            return response
        else:
            value='Please upload either excel or pdf only'
            return JsonResponse({'message': value},safe=False)

def process_reval_sgpa(request):
    supply_class_file= request.FILES.get('supply_class')
    supply_gpa_file= request.FILES.get("supply_gpa_class")
    file_mime_type = supply_class_file.content_type
    if file_mime_type == 'application/pdf':
        return_data=pdfToDataframe(supply_class_file)
        if isinstance(return_data, pd.DataFrame):
            reval_func(supply_gpa_file,return_data)
            response = FileResponse(open('Result.xlsx', 'rb'))
            response['Content-Disposition'] = 'attachment; filename="Result.xlsx"'
            return response
        elif isinstance(return_data,str):
            return JsonResponse({'message': return_data},safe=False)