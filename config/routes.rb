Rails.application.routes.draw do
  root 'home#index'
  get 'home/index'
  get 'home/preview'
  post 'home/parse_excel'
  get 'home/parse_excel'
  get 'home/print'
end
