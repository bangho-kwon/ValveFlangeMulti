# ValveFlangeMulti

Excel-driven Valve → Flange → (optional) Gasket placement UI.

## 빌드 요구사항

1. **ValveFlangeCore 프로젝트**: 이 프로젝트는 `ValveFlangeCore.csproj`를 참조합니다. 
   - 프로젝트 경로: `../ValveFlangeCore/ValveFlangeCore.csproj`
   - 이 프로젝트가 누락되면 컴파일 또는 런타임 오류가 발생합니다.

2. **Revit API**: REVIT_API_PATH 환경 변수 설정 필요
   - 기본값: `C:\Program Files\Autodesk\Revit 2023`

3. **.NET Framework 4.8**: 타겟 프레임워크

## 충돌 방지를 위한 개선사항

- 모든 주요 진입점에 예외 처리 추가
- Revit API 호출 전 null 체크 추가
- Excel 파일 로딩 오류 처리 강화
- 사용자에게 오류 정보 제공을 위한 상세 로깅
