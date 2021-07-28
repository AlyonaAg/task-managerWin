#include "Header.h"

// ����� ����� �� ����� � �� �������� � ������� ��������
void RecursiveTasksPrint(ITaskFolder* pCurFolder)
{
    HRESULT res;
    // �������� ������ �� ����� 
    IRegisteredTaskCollection* pTaskCollection = NULL;
    res = pCurFolder->GetTasks(NULL, &pTaskCollection);

    if (FAILED(res))
    {
        ERR("ITaskFolder::GetTasks", res);
        CoUninitialize();
        return;
    }

    LONG tasks_number = 0;
    res = pTaskCollection->get_Count(&tasks_number);

    if (tasks_number == 0)
    {
        pTaskCollection->Release();
        return;
    }

    TASK_STATE task_state;

    for (LONG i = 0; i < tasks_number; i++)
    {
        IRegisteredTask* pRegisteredTask = NULL;
        res = pTaskCollection->get_Item(_variant_t(i + 1), &pRegisteredTask);

        if (SUCCEEDED(res))
        {
            BSTR taskName = NULL;
            res = pRegisteredTask->get_Name(&taskName);
            if (SUCCEEDED(res))
            {
                printf("%d) %S \n", g_taskCount + 1, taskName);
                g_taskCount += 1;
                SysFreeString(taskName);

                res = pRegisteredTask->get_State(&task_state);
                if (SUCCEEDED(res))
                {
                    printf("\tState: %d \n", task_state);
                }
                else
                {
                    ERR("IRegisteredTask::get_State", res);
                }
            }
            else
            {
                ERR("IRegisteredTask::get_Name", res);
            }
            pRegisteredTask->Release();
        }
        else
        {
            ERR("IRegisteredTaskCollection::get_Item", res);
        }
    }

    pTaskCollection->Release();

    ITaskFolderCollection* pSubFolders = NULL;
    LONG subCount = 0;

    res = pCurFolder->GetFolders(0, &pSubFolders);
    if (FAILED(res))
    {
        ERR("ITaskFolder::GetFolders", res);
        return;
    }

    res = pSubFolders->get_Count(&subCount);
    if (FAILED(res))
    {
        ERR("ITaskFolderCollection::get_Count", res);
        pSubFolders->Release();
        return;
    }

    if (subCount != 0)
    {
        ITaskFolder* pSubFolder = NULL;
        for (LONG j = 0; j < subCount; ++j)
        {
            res = pSubFolders->get_Item(_variant_t(j + 1), &pSubFolder);

            if (SUCCEEDED(res))
            {
                RecursiveTasksPrint(pSubFolder);
                pSubFolder->Release();
            }
        }
    }
    pSubFolders->Release();
    return;
}

// ����� ���� �������� ����� 
void EnumerateActiveTasks()
{
    //  �������������� ��� 
    HRESULT res = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(res))
    {
        ERR("CoInitializeEx", res);
        return;
    }

    // ������������� ����� ������� ������������ ��� 
    res = CoInitializeSecurity(
        NULL,               // ��������� ���������� ������������
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,    // ������� ������������ 
        NULL,
        0,
        NULL);

    if (FAILED(res))
    {
        ERR("CoInitializeSecurity", res);
        CoUninitialize();
        return;
    }

    // ������� ��������� ������������ ����� ( Create an instance of the Task Service)  
    ITaskService* pService = NULL;
    res = CoCreateInstance(CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (void**)&pService);
    if (FAILED(res))
    {
        ERR("CoCreateInstance", res);
        CoUninitialize();
        return;
    }

    // ������������ � ������� ����� 
    res = pService->Connect(_variant_t(), _variant_t(),
        _variant_t(), _variant_t());
    if (FAILED(res))
    {
        ERR("ITaskService::Connect", res);
        pService->Release();
        CoUninitialize();
        return;
    }

    //�������� ��������� �� �������� ����� �����
    ITaskFolder* pRootFolder = NULL;
    res = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);

    pService->Release();
    if (FAILED(res))
    {
        ERR("ITaskService::GetFolder", res);
        CoUninitialize();
        return;
    }
    RecursiveTasksPrint(pRootFolder);
    pRootFolder->Release();
    CoUninitialize();
    return;
}

bool AddEventTrigger(ITriggerCollection* pTriggerCollection, wchar_t* szFilter)
{
    HRESULT res;
    //  �������� ������� �� ������� � ������.
    ITrigger* pTrigger = NULL;
    res = pTriggerCollection->Create(TASK_TRIGGER_EVENT, &pTrigger);
    if (FAILED(res))
    {
        ERR("ITriger::Create", res);
        CoUninitialize();
        return false;
    }

    IEventTrigger* pEventTrigger = NULL; //�������� ��������� �� ���������
    res = pTrigger->QueryInterface(
        IID_IEventTrigger, (void**)&pEventTrigger);
    pTrigger->Release();
    if (FAILED(res))
    {
        ERR("ITrigger::QueryInterface", res);
        CoUninitialize();
        return false;
    }

    // ���������� ���������� ����� �� 
    res = pEventTrigger->put_Delay(_bstr_t(L"PT0S"));

    // �������� ������ ������ �� ������ 
    res = pEventTrigger->put_Subscription(_bstr_t(szFilter));
    if (FAILED(res))
    {
        ERR("IEventTrigger::put_Subscription", res);
        pEventTrigger->Release();
        CoUninitialize();
        return false;
    }
    ITaskNamedValueCollection* pNamedValueQueries = NULL;
    res = pEventTrigger->get_ValueQueries(&pNamedValueQueries);
    pEventTrigger->Release();
    return true;
}

// ������� ������ � ����������� �� ������
void SetTask(enum Mode CurMode, LPWSTR szIpAddress)
{
    //  ������������� ��� 
    HRESULT res = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(res))
    {
        ERR("CoInitializeEx", res);
        return;
    }

    //  ���������� ����� ������� ������������ ���
    res = CoInitializeSecurity(
        NULL, //����� �������
        -1,   //���������� ������� � ��������� asAuthSvc
        NULL, //������ ����� ��������������
        NULL, 
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,//������� �������������� �� ��������� ��� �������� (���� ������)
        RPC_C_IMP_LEVEL_IMPERSONATE, //������� �������������
        NULL,
        0,
        NULL);

    if (FAILED(res))
    {
        ERR("CoInitializeSecurity", res);
        CoUninitialize();
        return;
    }

    // ������� ������ Task Service. 
    ITaskService* pService = NULL;
    res = CoCreateInstance(CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (void**)&pService);
    if (FAILED(res))
    {
        ERR("CoCreateInstance", res);
        CoUninitialize();
        return;
    }

    //  �����������  � ������� �����
    res = pService->Connect(_variant_t(), _variant_t(),
        _variant_t(), _variant_t());
    if (FAILED(res))
    {
        ERR("ITaskService::Connect", res);
        pService->Release();
        CoUninitialize();
        return;
    }

    // �������� ��������� �� �������� �����. � ��� ����� ���������������� ����� ������. 
    ITaskFolder* pRootFolder = NULL;
    res = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
    if (FAILED(res))
    {
        ERR("ITaskService::GetFolder", res);
        pService->Release();
        CoUninitialize();
        return;
    }

    // ���� ������ ��� ����������, ������� ��
    pRootFolder->DeleteTask(_bstr_t(wszTaskName), 0);

    // ������� ������ ITaskDefinition ��� �������� ������
    ITaskDefinition* pTask = NULL;
    res = pService->NewTask(0, &pTask);

    pService->Release();  // ������� ���
    if (FAILED(res))
    {
        ERR("ITaskService::NewTask", res);
        pRootFolder->Release();
        CoUninitialize();
        return;
    }

    //  �������� ��������� �� ������ ��������������� ����������
    IRegistrationInfo* pRegInfo = NULL;
    res = pTask->get_RegistrationInfo(&pRegInfo);
    if (FAILED(res))
    {
        ERR("ITaskDefinition::get_RegistrationInfo", res);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }

    res = pRegInfo->put_Author(_bstr_t(L"Ageeva Alyona"));
    if (FAILED(res))
    {
        ERR("IRegistrationInfo::put_Author", res);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }

    //  ������� ��������� ��� ������
    ITaskSettings* pSettings = NULL;
    res = pTask->get_Settings(&pSettings);
    if (FAILED(res))
    {
        ERR("ITaskDefinition::get_Setting", res);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }

    //  ��������� ��������� ������
    res = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
    pSettings->Release();
    if (FAILED(res))
    {
        ERR("ITaskSetting::put_StartWhenAvailable", res);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }

    //  �������� ��������� ���������, ����� �������� ������� �� �������.
    ITriggerCollection* pTriggerCollection = NULL;
    res = pTask->get_Triggers(&pTriggerCollection);
    if (FAILED(res))
    {
        ERR("ITaskDefinition::get_Triggers", res);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }

    if (CurMode == Ping_Mode)
    {
        if (!AddEventTrigger(pTriggerCollection, filterPing))
        {
            pRootFolder->Release();
            pTask->Release();
            pTriggerCollection->Release();
            CoUninitialize();
            return;
        }
    }
    else
    {
        if (!AddEventTrigger(pTriggerCollection, filterFirewall))
        {
            pRootFolder->Release();
            pTriggerCollection->Release();
            pTask->Release();
            CoUninitialize();
            return;
        }
        if (!AddEventTrigger(pTriggerCollection, filterWinSecurity))
        {
            pRootFolder->Release();
            pTriggerCollection->Release();
            pTask->Release();
            CoUninitialize();
            return;
        }
    }
    pTriggerCollection->Release();
    // ��� ������������ �������� ����� ������ ������ VB, ������� ������� MsgBox
    // � ������������
    IActionCollection* pActionCollection = NULL;

    // �������� ��������� �� ��������� ��������
    res = pTask->get_Actions(&pActionCollection);
    if (FAILED(res))
    {
        ERR("ITaskDefinition::get_Actions", res);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }

    // ����������� �������� �� ��������� ��������� 
    IAction* pAction = NULL;
    res = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
    pActionCollection->Release();
    if (FAILED(res))
    {
        ERR("IActionCollection::Create", res);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }

    IExecAction* pExecAction = NULL;
    res = pAction->QueryInterface(
        IID_IExecAction, (void**)&pExecAction);
    pAction->Release();
    if (FAILED(res))
    {
        ERR("IAction::GueryInterface", res);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }

    //  ����������� ���� � ������� 
    wchar_t* scriptPath = (CurMode == Ping_Mode) ? g_pingDetectedMsg : g_settingChangesMsg;
    res = pExecAction->put_Path(_bstr_t(scriptPath));
    pExecAction->Release();
    if (FAILED(res))
    {
        ERR("IExecAction::put_Path", res)
            pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }
    
    // ��� ������� ���������� � ping ����� ������������� �������� �������� 
    if (CurMode == Ping_Mode && szIpAddress != NULL)
    {
        pExecAction->put_Arguments(_bstr_t(szIpAddress));
    }
    // ��������� ������ � �������� ����� 
    // �� ����� ������ ������������
    IRegisteredTask* pRegisteredTask = NULL;

    res = pRootFolder->RegisterTaskDefinition(
        _bstr_t(wszTaskName),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(L"������������"),
        _variant_t(),
        TASK_LOGON_GROUP,
        _variant_t(L""),
        &pRegisteredTask);
    if (FAILED(res))
    {
        ERR("ITaskFolder::RegisterTaskDefinition", res);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return;
    }

    printf("\nSuccess! Task successfully registered.\n");

    // �������
    pRootFolder->Release();
    pTask->Release();
    pRegisteredTask->Release();
    CoUninitialize();
    return;
}

void PrintHelp() 
{
    wprintf(L"Usage:\nbsit_3_laba.exe [ping <source> | notice | clear]\n\
bsit_3_laba.exe - Enum all active tasks and their state\n\
bsit_3_laba.exe ping <source> - add a task to notice pings from a given IP\n\
bsit_3_laba.exe notice - add a task to notice changes in Windows Security settings\n\
bsit_3_laba.exe clear - delete previously created tasks\n");
}

void ClearTasks() {
    //  ������������� ��� 
    HRESULT res = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(res))
    {
        ERR("CoInitializeEx", res);
        return;
    }

    //  ���������� ����� ������� ������������ ���
    res = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL);

    if (FAILED(res))
    {
        ERR("CoInitializeSecurity", res);
        CoUninitialize();
        return;
    }

    // ������� ������ Task Service. 
    ITaskService* pService = NULL;
    res = CoCreateInstance(CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (void**)&pService);
    if (FAILED(res))
    {
        ERR("CoCreateInstance", res);
        CoUninitialize();
        return;
    }

    //  �����������  � ������� �����
    res = pService->Connect(_variant_t(), _variant_t(),
        _variant_t(), _variant_t());
    if (FAILED(res))
    {
        ERR("ITaskService::Connect", res);
        pService->Release();
        CoUninitialize();
        return;
    }

    // �������� ��������� �� �������� �����
    ITaskFolder* pRootFolder = NULL;
    res = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
    if (FAILED(res))
    {
        ERR("ITaskService::GetFolder", res);
        pService->Release();
        CoUninitialize();
        return;
    }

    // ���� ������ ��� ����������, ������� ��
    res = pRootFolder->DeleteTask(_bstr_t(wszTaskName), 0);
    if (res == S_OK) printf("\nTask successfully deleted!");
    // �������
    pRootFolder->Release();
    CoUninitialize();
    return;
}

void PingTask(LPWSTR szIpAddress) {
    wchar_t* tempStr = (wchar_t*)calloc(256, sizeof(wchar_t));
    if (tempStr != NULL)
    {
        wcscpy(tempStr, filterPing);
        wsprintfW(filterPing, tempStr, szIpAddress);
        free(tempStr);
        SetTask(Ping_Mode, szIpAddress);
    }
}

int wmain(int argc, wchar_t* argv[])
{
    setlocale(LC_ALL, "Rus");
    if (argc == 1)
    {
        EnumerateActiveTasks();
    }
    else
    {
        //��������� ������ ��� ����� � �������
        WCHAR szCurDirectory[BUFSIZ];
        GetCurrentDirectoryW(BUFSIZ, szCurDirectory);

        g_pingDetectedMsg = (LPWSTR)calloc(wcslen(szCurDirectory) + 32, sizeof(WCHAR));
        g_settingChangesMsg = (LPWSTR)calloc(wcslen(szCurDirectory) + 32, sizeof(WCHAR));
        if (g_pingDetectedMsg != NULL
            && g_settingChangesMsg != NULL)
        {
            wcscpy(g_pingDetectedMsg, szCurDirectory);
            wcscat(g_pingDetectedMsg, PingMsgName);

            wcscpy(g_settingChangesMsg, szCurDirectory);
            wcscat(g_settingChangesMsg, ChangesMsgName);
        }
        //����� � ����������� �� ������� ����������
        if (!wcscmp(L"clear", argv[1])) ClearTasks();
        else if (!wcscmp(L"help", argv[1])) PrintHelp();
        else if (!wcscmp(L"ping", argv[1]) && argc > 2) PingTask(argv[2]);
        else if (!wcscmp(L"notice", argv[1])) SetTask(Update_Mode, NULL);
        else printf("Wrong input!");

        //������ ������ ����� ������ 
        if (g_settingChangesMsg != NULL)
        {
            free(g_settingChangesMsg);
            g_settingChangesMsg = NULL;
        }
        if (g_pingDetectedMsg != NULL)
        {
            free(g_pingDetectedMsg);
            g_pingDetectedMsg = NULL;
        }
    }
    return 0;
}
