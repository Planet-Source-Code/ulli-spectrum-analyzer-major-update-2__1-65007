
    use32

    xor     eax, eax
    mov     edx, [esp + 8]
    mov     ecx, [esp + 12]
    jcxz    GetOut

Again:
    rcr     edx, 1
    adc     eax, eax
    dec     ecx
    jnz     Again

GetOut:
    mov     edx, [esp + 16]
    mov     [edx], eax
    ret     16