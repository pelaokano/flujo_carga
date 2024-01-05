#include <iostream>
#include <Windows.h>
#include <vector>
#include <string>
#include <time.h>
#include <Api.hpp>
#include <ExitError.hpp>

typedef api::v2::DataObject DataObject;
typedef api::v2::Application Application;
typedef api::v2::Api Api;
typedef api::v2::ExitError ExitError;

using api::Value;
using std::vector;
using std::string;

int main()
{
    // Se carga la DLL con los metodos de la API de c++ de Digsilent
    HINSTANCE dllHandle = NULL;
    dllHandle = LoadLibraryEx(TEXT("C:\\Program Files\\DIgSILENT\\PowerFactory 2022 SP4\\digapi.dll"), NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
    if (!dllHandle) {
        std::cout << "Error al cargar dll" << std::endl;
        return 1;
    }
    else {
        std::cout << "Dll cargada exitosamente" << std::endl;
    }

    CREATEAPI_V2 createApi = (CREATEAPI_V2)GetProcAddress(dllHandle, "CreateApiInstanceV2");
    DESTROYAPI_V2 destroyApi = (DESTROYAPI_V2)GetProcAddress(dllHandle, "DestroyApiInstanceV2");

    Api* apiInstance = createApi(NULL, NULL, NULL);
    int error = 0;

    if (apiInstance) {
        std::cout << "Se crea la instancia de la api" << std::endl;
        // destroyApi(apiInstance);
        Application* app = apiInstance->GetApplication();
        DataObject* user = app->GetCurrentUser();
        std::cout << "Nombre del usuario: " << user->GetName()->GetString() << std::endl;
        std::string path_project = "39 Bus New England System.IntPrj";
        Value rutaProyecto(path_project.c_str());
        const Value* proyectoValue = user->Execute("SearchObject", &rutaProyecto, &error);
        DataObject* proyecto = static_cast<DataObject*> (proyectoValue->GetDataObject());
        proyecto->Execute("Activate", NULL, &error);

        Value comLdfString("ComLdf");
        const Value* comLdfValue = app->Execute("GetCaseObject", &comLdfString);
        DataObject* comLdfObj = static_cast<DataObject*>(comLdfValue->GetDataObject());
        comLdfObj->SetAttributeInt("e:iopt_net", 1, &error);
        comLdfObj->SetAttributeInt("e:iopt_plim", 0, &error);
        comLdfObj->SetAttributeInt("e:iopt_tem", 0, &error);
        comLdfObj->SetAttributeDouble("e:errlf", 1, &error);
        const Value* executeLdf = comLdfObj->Execute("Execute", NULL, &error);

        Value args(Value::VECTOR);
        args.VecInsertString("*.ElmLne,*.ElmTr2");
        args.VecInsertInteger(0);
        const Value* ElemInService = app->Execute("GetCalcRelevantObjects", &args, &error);

        if (error == 0) {
            for (unsigned int i = 0; i < ElemInService->VecGetSize(); i++) {
                DataObject* elemObject = static_cast<DataObject*>(ElemInService->VecGetDataObject(i));
                const Value* nameElem = elemObject->GetName();
                double loadingElem = elemObject->GetAttributeDouble("c:loading", &error);
                
                std::cout << "nombre elemento: " << nameElem->GetString() << ", loading: " << loadingElem << std::endl;

                apiInstance->ReleaseValue(nameElem);

            }
        }
        apiInstance->ReleaseValue(ElemInService);
        apiInstance->ReleaseValue(executeLdf);
        apiInstance->ReleaseValue(proyectoValue);
        apiInstance->ReleaseValue(comLdfValue);
        
        apiInstance->ReleaseObject(comLdfObj);
        apiInstance->ReleaseObject(proyecto);
        apiInstance->ReleaseObject(user);

    }
    else {
        std::cout << "Error al crear la instancia de la api" << std::endl;
    }
    destroyApi(apiInstance);
    FreeLibrary(dllHandle);
    return 0;
}