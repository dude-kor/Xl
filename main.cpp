#include "CXl.h"

int main() {
    CXl xl;
    xl.Open();

    //printf("SetSafeBound %d\n", xl.SetSafeBound(1,20,1,26));
    printf("SetSafeBound %d\n", xl.SetSafeBound("A1", "Z26"));

    //���⼭���� ���� �۾��� ���ֽø� �˴ϴ�.
    printf("SetData %d\n", xl.SetData("hi", 1, 1));

    printf("SetData2 %d\n", xl.SetData("hello", 1, 2));
    //��������� �۾��Ͻø� �˴ϴ�.

    printf("%d\n", xl.AddActiveSheet());

    printf("%d\n", xl.SetRange("A1", "Z15"));

    printf("%d\n", xl.Print());

    //printf("%d", xl.Save());


    //xl.Close();

    //CoUninitialize();
    return 0;
}