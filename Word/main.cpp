#include <iostream>
#include <time.h>
#include <Windows.h>

using namespace std;

int main()
{
    char wrd;
    cout << "Press A to start" << endl;
    cin >> wrd;
    if(wrd=='a')
    {
        system("start winword");
    }
    return 0;
}
