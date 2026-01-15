# SCVMM-log-collection

(Korean)
.\Get-VmmDiag.ps1 -RefreshLLDP -JobHistoryHours 24 -Credential (Get-Credential)
출력: C:\VMMReports\SCVMM_Diag_<timestamp>\ 폴더에 HTML 보고서 + CSV + 호스트별 WinRM 설정 텍스트가 생성됩니다.
레거시 **NetBIOS(139/tcp)**도 테스트하려면 -IncludeLegacyNetBIOS를 추가하세요.

(English)
.\Get-VmmDiag.ps1 -RefreshLLDP -JobHistoryHours 24 -Credential (Get-Credential)
The output will be generated in the folder: C:\VMMReports\SCVMM_Diag_<timestamp>\ and will include an HTML report, CSV files, and per-host WinRM configuration text files.
To also test legacy NetBIOS (139/tcp), add the -IncludeLegacyNetBIOS parameter.
