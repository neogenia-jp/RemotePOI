#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
USER app
WORKDIR /app
EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["XlsManiSvc/XlsManiSvc.csproj", "XlsManiSvc/"]
RUN dotnet restore "XlsManiSvc/XlsManiSvc.csproj"
COPY . .
WORKDIR "/src/XlsManiSvc"
RUN dotnet build "XlsManiSvc.csproj" -c $BUILD_CONFIGURATION -o /app/build

FROM build AS publish
RUN dotnet publish "XlsManiSvc.csproj" -c $BUILD_CONFIGURATION -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "XlsManiSvc.dll"]