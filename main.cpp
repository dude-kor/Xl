<<<<<<< HEAD
#include "CXl.h"

int main() {
    CXl xl;
    xl.Open();

    //printf("SetSafeBound %d\n", xl.SetSafeBound(1,20,1,26));
    printf("SetSafeBound %d\n", xl.SetSafeBound("A1", "Z26"));

    //여기서부터 엑셀 작업을 해주시면 됩니다.
    printf("SetData %d\n", xl.SetData("hi", 1, 1));

    printf("SetData2 %d\n", xl.SetData("hello", 1, 2));
    //여기까지만 작업하시면 됩니다.

    printf("%d\n", xl.AddActiveSheet());

    printf("%d\n", xl.SetRange("A1", "Z15"));

    printf("%d\n", xl.Print());

    //printf("%d", xl.Save());


    //xl.Close();

    //CoUninitialize();
    return 0;
=======
#include "CXl.h"

int main() {
    CXl xl;
    xl.Open();

    //printf("SetSafeBound %d\n", xl.SetSafeBound(1,20,1,26));
    printf("SetSafeBound %d\n", xl.SetSafeBound("A1", "Z26"));

    //여기서부터 엑셀 작업을 해주시면 됩니다.
    printf("SetData %d\n", xl.SetData("hi", 1, 1));

    printf("SetData2 %d\n", xl.SetData("hello", 1, 2));
    //여기까지만 작업하시면 됩니다.

    printf("%d\n", xl.AddActiveSheet());

    printf("%d\n", xl.SetRange("A1", "Z15"));

    printf("%d\n", xl.Print());

    //printf("%d", xl.Save());


    //xl.Close();

    //CoUninitialize();
    return 0;
>>>>>>> 2b52f65a308d6593fa542001cff22d180e17066f
}